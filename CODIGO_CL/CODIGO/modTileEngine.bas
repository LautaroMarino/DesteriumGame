Attribute VB_Name = "Mod_TileEngine"
' Externo
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public MaxGrh As Long

Private Const TARGET_FPS_MS As Long = 1000

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (lpvDest As Any, _
                                      lpvSource As Any, _
                                      ByVal cbCopy As Long)

Public vSyncRun       As Boolean
Public AmbientColor   As RGBA

Public g_Last_OffsetX As Single
Public g_Last_OffsetY As Single

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 _
                Lib "gdi32" _
                Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                               ByVal lpsz As String, _
                                               ByVal cbString As Long, _
                                               lpSize As Size) As Long

Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

' Interno
' Exodo Online 0.13.5
' #Include Wgl_Client.dll

'Techos
Public Const ROOF_ALPHA_SPEED = 0.3
Public Const ROOF_ALPHA_MAX As Byte = 255
Public Const ROOF_ALPHA_MIN As Byte = 0
Public RoofAlpha As Single

' CONSTANT GENERAL
Public Const INFINITE_LOOPS As Integer = -1

Private Const GrhFogata      As Long = 1521

Private Const GrhFogata2      As Long = 70272

' Colores de npc
Public Enum eNPCType

    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    Guardiascaos = 8
    ResucitadorNewbie = 9
    Pretoriano = 10
    Gobernador = 11
    Mascota = 12
    Fundition = 13
    eYacimiento = 14        ' Detecta si es un Yacimiento (No Ataca)

End Enum

' Directorios
Public IniPath                 As String

Public MapPath                 As String

'Bordes del mapa
Public MinXBorder              As Byte

Public MaxXBorder              As Byte

Public MinYBorder              As Byte

Public MaxYBorder              As Byte

'Status del user
Public CurMap                  As Integer 'Mapa actual

Public UserIndex               As Integer
'Public WaitInput As Boolean
Public UserMoving              As Boolean

Public UserBody                As Integer

Public UserHead                As Integer

Public UserPos                 As Position 'Posicion

Public AddtoUserPos            As Position 'Si se mueve

Public UserCharIndex           As Integer

' Engine andando
Public EngineRun               As Boolean

' Time FPS
Public lFrameTimer             As Long
Public FrameTime             As Long
Public FPS                     As Long

Public UserFps As Long


Private fpsLastCheck           As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth        As Integer

Public WindowTileHeight       As Integer

Public HalfWindowTileWidth    As Integer

Public HalfWindowTileHeight   As Integer

'Offset del desde 0,0 del main view

Public OffsetCounterX         As Single

Public OffsetCounterY         As Single

'Tamaño de los tiles en pixels
Public TilePixelHeight         As Integer

Public TilePixelWidth          As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX   As Single

Public ScrollPixelsPerFrameY   As Single

Private timerEngine            As Currency
Dim timerElapsedTime           As Single
Public timerTicksPerFrame         As Single
Dim engineBaseSpeed            As Single

Public NumBodies               As Integer

Public Numheads                As Integer

Public NumFxs                  As Integer

Public NUMCHARS                As Integer

Public LastChar                As Integer

Public NumWeaponAnims          As Integer

Public NumShieldAnims          As Integer

Private MouseTileX             As Byte

Private MouseTileY             As Byte

Public MapInfo                 As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bTecho                  As Boolean 'hay techo?

Public brstTick                As Long

' Used by GetTextExtentPoint32
Private Type Size

    cx As Long
    cy As Long

End Type

' Wgl_Client.DLL
' CONSTANT


Public MAX_TEXTURE               As Long

Public Const MAX_TEXTURE_GUI     As Long = 500

Public Const MAX_TEXTURE_AVATARS    As Long = 157

Public Const MAX_CAPTION         As Long = 52

Public Const MAX_FONT            As Long = 6

Public Const MAX_TECHNIQUE       As Long = 1

Public g_Textures()              As Integer
Public g_Textures_Gui()              As Integer
Public g_Textures_MiniMapa()              As Integer
Public g_Textures_Avatars()              As Integer

Public g_Captions(MAX_CAPTION)   As Integer

Public Fonts(MAX_FONT)           As Integer

Public Techniques(MAX_TECHNIQUE) As Integer

' QuadTree (wolfy lindo)
Public g_Swarm                   As New wGL_Temp_Swarm

' Modo Noche
Public ModoNoche                 As Boolean

' Ventanas disponibles
Public Enum eCaption

    MainPicture = 0 ' Ventana de inventario default
    CuentaPicture = 1 ' Ventana de la cuenta (Personajes dibujados de cuenta)
    
    Comercio_User = 2
    Comercio_Npc = 3
    
    Boveda_Npc = 4
    Boveda_User = 5
    
    RankTop = 6
    eInvRuleta = 7
    
    cInvComUsu = 8
    cInvOfferComUsu1 = 9
    cInvOfferComUsu2 = 10
    cInvOroComUsu1 = 11
    cInvOroComUsu2 = 12
    cInvOroComUsu3 = 13
    cSubastar = 14
    cPicEvent = 15 ' Inventario de los Objetos Donados
    cPicEventChar = 16 ' Char con el objeto equipado o visualizado en caso de no ser equipable.
    cPivQuest = 17 ' Quest de cada NPC.
    cCriaturaInfo = 18 ' Form con informacion de la criatura seleccionada.
    cObjectInfo = 19   ' Form con informacion del objecto seleccionado.
    cMiniMapa = 20
    cMapaGrande = 21
    cMercader = 22
    cMercaderList = 23
    cMercaderInv = 24
    cStats = 25
    cMapGrande = 26




    cInvCofre1 = 31
    cInvCofre2 = 32
    cInvMapa = 33
    
    cInvEldhirComUsu1 = 34
    cInvEldhirComUsu2 = 35
    cInvEldhirComUsu3 = 36
    
    eCommerceChar = 37
    eInvSkin1 = 38
    eguildpanel = 39
     eInvSkin2 = 40
    eMapa = 41
    eMapaNpc = 42
    eMapaObj = 43
    
    eCharAccount = 44
    
    eMercader_Inv = 45
    eMercader_Bank = 46
    eMercader_List = 47
    eMercader_Meditations = 48
    
    eMercader_ListOffer = 49
    
    e_Shop = 50
    
    e_Perfil = 51
    e_Objetivos = 52
End Enum

' Fuentes disponibles
Public Enum eFonts

    f_Tahoma = 0
    f_Morpheus = 1
    f_Chat = 2
    f_Emoji = 3
    f_Medieval = 4
    f_Verdana = 5
    f_Booter = 6

End Enum

' Técnicas disponibles
Public Enum eTechnique

    t_Default = 0
    t_Alpha = 1

End Enum

Public Sub Set_engineBaseSpeed(ByVal Value As Single)
    engineBaseSpeed = Value
End Sub
Public Function InitTileEngine(ByVal setTilePixelHeight As Integer, _
                               ByVal setTilePixelWidth As Integer, _
                               ByVal setWindowTileHeight As Integer, _
                               ByVal setWindowTileWidth As Integer, _
                               ByVal pixelsToScrollPerFrameX As Single, _
                               ByVal pixelsToScrollPerFrameY As Single, _
                               ByVal engineSpeed As Single) As Boolean
        '<EhHeader>
        On Error GoTo InitTileEngine_Err
        '</EhHeader>
                                
        '***************************************************
        'Author: WAICON
        'Last Modification: 05-05-2019
        '
        'Creates all DX objects and configures the engine to start running.
        '***************************************************
    
        'Fill startup variables
100     TilePixelWidth = setTilePixelWidth
102     TilePixelHeight = setTilePixelHeight
104     WindowTileHeight = setWindowTileHeight
106     WindowTileWidth = setWindowTileWidth
    
108     HalfWindowTileHeight = setWindowTileHeight \ 2
110     HalfWindowTileWidth = setWindowTileWidth \ 2
    
112     engineBaseSpeed = engineSpeed
    
        'Set FPS value to 60 for startup
114     FPS = 60
    
118     MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
120     MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
122     MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
124     MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
        'Resize mapdata array
130     ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
132     ReDim MapData_Copy(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
  
        'Set intial user position
134     UserPos.X = MinXBorder
136     UserPos.Y = MinYBorder
    
        'Set scroll pixels per frame
138     ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
140     ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

        Dim Engine As Boolean

        Dim Modo   As wGL_Graphic_Mode
    
142     If ClientSetup.bFps = 2 Then
144         Modo = MODE_SYNCHRONISED Or MODE_COMPATIBLE

        Else
146         Modo = MODE_COMPATIBLE

        End If
    
    
        Dim Mult As Byte
        
        #If ModoBig = 0 Then
148     Mult = 1
        #Else
          Mult = 2
        #End If
            
            
            Dim A As Long
            

             Engine = wGL_Graphic.Create_Driver(DRIVER_DIRECT3D9, Modo, FrmMain.MainViewPic.hWnd, FrmMain.MainViewPic.ScaleWidth, FrmMain.MainViewPic.ScaleHeight)

        
         
150     If (Not Engine) Then
152         MsgBox "Parece ser que su PC no es compatible con Aurora.Multimedia."
154         CloseClient
        End If
    
156     Call Load_Technique
158     Call Load_Font
160     Call Load_Captions
    
162     InitTileEngine = True
        '<EhFooter>
        Exit Function

InitTileEngine_Err:
        LogError err.Description & vbCrLf & _
               "in InitTileEngine " & _
               "at line " & Erl

        '</EhFooter>
End Function

Public Sub Draw_InventoryCaption(ByVal GrhIndex As Long, _
                                 ByVal CaptionIndex As Integer, _
                                 ByVal PicWidth As Integer, _
                                 ByVal PicHeight As Integer, _
                                 ByVal Width As Integer, _
                                 ByVal Height As Integer, _
                                 ByVal Colour As Long)
          
    Call wGL_Graphic.Use_Device(g_Captions(CaptionIndex))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, Width, Height)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    Call Draw_Texture(GrhIndex, 1, 1, To_Depth(1), Width, Height, Colour, 0, eTechnique.t_Default)
    
    Call wGL_Graphic_Renderer.Flush
    
End Sub

Private Sub Load_Captions()
    ' Seteamos todas las ventanas que tengamos
    
    Dim Mult As Byte
    Dim Width As Integer, Height As Integer
    
    ' Inventario
    Set Inventario = New clsGrapchicalInventory
    
    ' Ventana principal de inventario
    #If ModoBig > 0 Then
        Mult = 1
        Width = 70
        Height = 70
        
        g_Captions(eCaption.e_Perfil) = wGL_Graphic.Create_Device_From_Display(FrmMain.PicPerfil.hWnd, FrmMain.PicPerfil.ScaleWidth, FrmMain.PicPerfil.ScaleHeight)
    #Else
        Mult = 1
        Width = 35
        Height = 35
    #End If
    
    g_Captions(eCaption.MainPicture) = wGL_Graphic.Create_Device_From_Display(FrmMain.PicInv.hWnd, FrmMain.PicInv.ScaleWidth, FrmMain.PicInv.ScaleHeight)
    Call Inventario.Initialize(FrmMain.PicInv, MAX_INVENTORY_SLOTS, MAX_INVENTORY_SLOTS, eCaption.MainPicture, Width, Height, , , , , , True, , , True, True)
    g_Captions(eCaption.cMiniMapa) = wGL_Graphic.Create_Device_From_Display(FrmMain.MiniMapa.hWnd, FrmMain.MiniMapa.ScaleWidth, FrmMain.MiniMapa.ScaleHeight)

    
    'Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cMiniMapa))
    'Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.MainPicture))
    
End Sub

Private Sub Load_Font()

    Const filePath As String = "\resource\FONT\"
    
    ' esto es para el cliente
    ' pero todavia no, le falta testeo a despues te lo explico otro dia xD
    'joya lo espero ! tranqui  y genialk el efecto este de fuentes :D salgo
    ' Cargado de fuentes
    Fonts(eFonts.f_Tahoma) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "Tahoma.TTF"))
    Fonts(eFonts.f_Morpheus) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "Morpheus.TTF"))
    Fonts(eFonts.f_Chat) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "tahomabd.ttf"))
    Fonts(eFonts.f_Emoji) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "TwemojiMozilla.ttf"))
    Fonts(eFonts.f_Medieval) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "Medieval.ttf"))
    Fonts(eFonts.f_Verdana) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "Verdana.ttf"))
    
    
    ' Todas las fuentes excepto emoji van a chequear los emoji en f_Emoji
    Call wGL_Graphic_Renderer.Update_Font_Fallback(Fonts(eFonts.f_Chat), Fonts(eFonts.f_Emoji))
    'Call wGL_Graphic_Renderer.Update_Font_Fallback(Fonts(eFonts.f_Morpheus), Fonts(eFonts.f_Emoji))
    
    
    'Call AddFont(App.path & FilePath & "booterfz.ttf")
    Fonts(eFonts.f_Booter) = wGL_Graphic_Renderer.Create_Font(LoadBytes(filePath & "booterfz2.ttf"))
End Sub

Private Sub Load_Technique()
    ' Cargado de técnicas
    
    Dim Descriptor As wGL_Graphic_Descriptor

    Descriptor.Depth = COMPARISON_LESS_EQUAL
    Descriptor.Depth_Mask = True
    Descriptor.Mask_Red = True: Descriptor.Mask_Green = True: Descriptor.Mask_Blue = True: Descriptor.Mask_Alpha = True
    Descriptor.Stencil_Mask = &HFF
    
    Techniques(eTechnique.t_Default) = wGL_Graphic_Renderer.Create_Technique

    Dim Program As Integer

    Program = wGL_Graphic.Create_Program(LoadBytes("\resource\shader\VsSprite.bin"), LoadBytes("\resource\shader\PsSpriteOpaque.bin"))
    
    Call wGL_Graphic_Renderer.Update_Technique_Descriptor(Techniques(eTechnique.t_Default), Descriptor)
    Call wGL_Graphic_Renderer.Update_Technique_Program(Techniques(eTechnique.t_Default), Program)
        
    Techniques(eTechnique.t_Alpha) = wGL_Graphic_Renderer.Create_Technique
    Program = wGL_Graphic.Create_Program(LoadBytes("\resource\shader\VsSprite.bin"), LoadBytes("\resource\shader\PsSpriteTransparent.bin"))

    Descriptor.Depth_Mask = False
    Descriptor.Blend_Color_Source = BLEND_FACTOR_SRC_ALPHA
    Descriptor.Blend_Color_Destination = BLEND_FACTOR_ONE_MINUS_SRC_ALPHA
    
    Call wGL_Graphic_Renderer.Update_Technique_Descriptor(Techniques(eTechnique.t_Alpha), Descriptor)
    Call wGL_Graphic_Renderer.Update_Technique_Program(Techniques(eTechnique.t_Alpha), Program)
    
    ' Inityial state
    Call SetAmbientColor(255, 255, 255, 255)

End Sub

Public Sub DeinitTileEngine()

    On Error Resume Next

    Set InvComUser = Nothing
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvEldhirComUsu(0) = Nothing
    Set InvEldhirComUsu(1) = Nothing
    Set InvEldhirComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing

End Sub

Public Sub RenderStarted()

    Dim Value As Long
    
    'Sólo dibujamos si la ventana no está minimizada
    If (FrmMain.WindowState <> 1 And FrmMain.visible) Or MirandoCuenta Then

        If ShowNextFrame() Then
            FPS = FPS + 1
        End If
        
        If FrmMain.PicInv.visible Then
            Call Inventario.DrawInventory

        End If
        
       ' Value = Value + 1
        
       ' If Value = 10 Then
        'Sleep 1
           ' Value = 0
      '  End If
    Else
        Sleep 1
    End If
    
    DoEvents
    Call wGL_Graphic.Commit

    Call CheckKeys
    Moviendose = False
    
    'FPS Counter - mostramos las FPS
    If FrameTime - lFrameTimer >= 1000 Then
        UserFps = FPS
        FrmMain.lblFPS.Caption = UserFps
        lFrameTimer = FrameTime
        FPS = 0

    End If

End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef tX As Byte, _
                  ByRef tY As Byte)
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
    On Error Resume Next
    'tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    'tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
    'tX = (UserPos.X - HalfWindowTileWidth) + viewPortX \ TilePixelWidth
    'tY = (UserPos.Y - HalfWindowTileHeight) + viewPortY \ TilePixelHeight
    
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
    
    If tX <= 0 Then tX = 1
    If tY <= 0 Then tY = 1
    
End Sub

Sub MakeChar(ByVal CharIndex As Integer, _
             ByVal Body As Integer, _
             ByVal BodyAttack As Integer, _
             ByVal Head As Integer, _
             ByVal Heading As Byte, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Arma As Integer, _
             ByVal Escudo As Integer, _
             ByVal Casco As Integer, _
             ByRef AuraIndex() As Byte, _
             ByVal NpcIndex As Integer)

        '<EhHeader>
        On Error GoTo MakeChar_Err

        '</EhHeader>

        'Apuntamos al ultimo Char
100     If CharIndex > LastChar Then LastChar = CharIndex
    
102     With CharList(CharIndex)

            'If the char wasn't allready active (we are rewritting it) don't increase char count
104         'If .Active = 0 Then NUMCHARS = NUMCHARS + 1
        
106         If Arma = 0 Then Arma = 2
108         If Escudo = 0 Then Escudo = 2
110         If Casco = 0 Then Casco = 2
112
        
114         .iHead = Head
116         .iBody = Body
118         .Head = HeadData(Head)
120         .Body = BodyData(Body)
122         .BodyAttack = BodyDataAttack(BodyAttack)
124         .Arma = WeaponAnimData(Arma)

            Dim A As Long

            For A = 1 To MAX_AURAS

                If AuraIndex(A) = 0 Then AuraIndex(A) = 2
126             .Aura(A) = AuraAnimData(AuraIndex(A))
            Next A
              
128         .Escudo = ShieldAnimData(Escudo)
130         .Casco = CascoAnimData(Casco)
            .NpcIndex = NpcIndex
            
132         If Casco = 54 Then
134             .OffsetY = -15
            Else
136             .OffsetY = 0

            End If
        
138         .Heading = Heading
        
            'Reset moving stats
            If Not CharIndex = UserCharIndex Then
                .Moving = False
                .MoveOffsetX = 0
                .MoveOffsetY = 0

            End If
        
            'Reset moving stats
140         If (CharIndex = UserCharIndex) Then
142             .MoveOffsetX = g_Last_OffsetX
144             .MoveOffsetY = g_Last_OffsetY
146             g_Last_OffsetX = 0
148             g_Last_OffsetY = 0

            End If

            'Update position
166         .Pos.X = X
168         .Pos.Y = Y
                    
            ' Create virtual sound source
            .SoundSource = Audio.CreateSource(X, Y)
            
            'Make active
170         .Active = 1
        
            ' Update QuadTree
            Dim RangeX As Single, RangeY As Single

172         Call GetCharacterDimension(CharIndex, RangeX, RangeY)
174         Call g_Swarm.Insert(5, CharIndex, X, Y, RangeX, RangeY)
        
        
        End With
    
        'Plot on map
176     MapData(X, Y).CharIndex = CharIndex

        '<EhFooter>
        Exit Sub

MakeChar_Err:
        LogError err.Description & vbCrLf & "in MakeChar " & "at line " & Erl

        '</EhFooter>
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    With CharList(CharIndex)
        .Speeding = 0
        .LastStep = 0
        .Idle = False
        .Nombre = vbNullString
        .LastDialog = vbNullString
        .GuildName = vbNullString
        .Active = 0
        .Criminal = 0
        .FxIndex = 0
        .Moving = False
        .Pos.X = 0
        .Pos.Y = 0
        
        .Pie = False
        .UsandoArma = False
        .TimeAttackNpc = 0
        .Muerto = False
        .Invisible = False
        .Intermitencia = False
        .Atacable = False
        .NpcIndex = 0
        .ColorNick = 0
        .GroupIndex = 0
        
        .MinHp = 0
        .MaxHp = 0
        .MinMan = 0
        .MaxMan = 0
    End With
    
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
        '<EhHeader>
        On Error GoTo EraseChar_Err
        '</EhHeader>

        '*****************************************************************
        'Erases a character from CharList and map
        '*****************************************************************

100     CharList(CharIndex).Active = 0
    
        'Update lastchar
102     If CharIndex = LastChar Then

104         Do Until CharList(LastChar).Active = 1
106             LastChar = LastChar - 1

108             If LastChar = 0 Then Exit Do
            Loop

        End If
    
110     If CharList(CharIndex).Pos.X > 0 Then
112         MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0
        End If
    
114     Call g_Swarm.Remove(5, CharIndex, 0, 0, 0, 0)
    
        'Remove char's dialog
116     Call Dialogos.RemoveDialog(CharIndex)
    
118     Call ResetCharInfo(CharIndex)
                
        ' Destroy virtual sound source
        Call Audio.DeleteSource(CharList(CharIndex).SoundSource, False)
            
        'Update NumChars
120     'NUMCHARS = NUMCHARS - 1
        '<EhFooter>
        Exit Sub

EraseChar_Err:
        LogError err.Description & vbCrLf & _
           "in EraseChar " & _
           "at line " & Erl

        '</EhFooter>
End Sub

Public Sub InitGrh(ByRef grh As grh, _
                   ByVal GrhIndex As Long, _
                   Optional ByVal started As Long = -1, _
                   Optional ByVal Loops As Integer = INFINITE_LOOPS)
        '<EhHeader>
        On Error GoTo InitGrh_Err
        '</EhHeader>

100     If GrhIndex = 0 Or GrhIndex > MaxGrh Then Exit Sub
    
102     grh.GrhIndex = GrhIndex

104     If GrhData(GrhIndex).NumFrames > 1 Then
106         If started >= 0 Then
108             grh.started = started
            Else
110             grh.started = FrameTime

            End If
        
112         grh.Loops = Loops
114         grh.Speed = GrhData(GrhIndex).Speed / GrhData(GrhIndex).NumFrames


        Else
116         grh.started = 0
118         grh.Speed = 1

        End If
    
        '<EhFooter>
        Exit Sub

InitGrh_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Mod_TileEngine.InitGrh " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
        '*****************************************************************
        
        'Starts the movement of a character in nHeading direction
        '*****************************************************************
        '<EhHeader>
        On Error GoTo MoveCharbyHead_Err

        '</EhHeader>
        Dim AddX As Integer

        Dim AddY As Integer

        Dim X    As Integer

        Dim Y    As Integer

        Dim nX   As Integer

        Dim nY   As Integer
    
100     With CharList(CharIndex)
102         X = .Pos.X
104         Y = .Pos.Y
        
            'Figure out which way to move
106         Select Case nHeading

                Case E_Heading.NORTH
108                 AddY = -1
        
110             Case E_Heading.EAST
112                 AddX = 1
        
114             Case E_Heading.SOUTH
116                 AddY = 1
            
118             Case E_Heading.WEST
120                 AddX = -1

            End Select
        
122         nX = X + AddX
124         nY = Y + AddY
                
126         MapData(nX, nY).CharIndex = CharIndex
128         .Pos.X = nX
130         .Pos.Y = nY
        
132         If (MapData(X, Y).CharIndex = CharIndex) Then
134             MapData(X, Y).CharIndex = 0

            End If
        
136         Call g_Swarm.Move(CharIndex, nX, nY)
            
138         .MoveOffsetX = -1 * (TilePixelWidth * AddX)
140         .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
            Call Audio.UpdateSource(.SoundSource, nX, nY)

144         .Heading = nHeading
            
            ' Sistema de Escaleras 65562 65563
            If MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).Graphic(2).GrhIndex = 65562 Or MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).Graphic(2).GrhIndex = 65563 Then
                
                CharList(CharIndex).Heading = NORTH

            End If

156         .scrollDirectionX = AddX
158         .scrollDirectionY = AddY

            Draw_MiniMap

            .Idle = False
            
            If Not .Moving Then

                'If .Muerto Then
                    '.Body = BodyData(CASPER_BODY)

                'End If

                'Start animations
                If .Body.Walk(.Heading).started = 0 Then
                    .Body.Walk(.Heading).started = FrameTime
                    .Arma.WeaponWalk(.Heading).started = FrameTime
                    .Escudo.ShieldWalk(.Heading).started = FrameTime

                    .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                    .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS

                End If
            
                .MovArmaEscudo = False
                .Moving = True

            End If

        End With
    
160     If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
        '<EhFooter>
        Exit Sub

MoveCharbyHead_Err:
        LogError err.Description & vbCrLf & "in MoveCharbyHead " & "at line " & Erl

        '</EhFooter>
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 09/21/2010
    ' 09/21/2010: C4b3z0n - Changed from Private Funtion tu Public Function.
    '***************************************************
    With CharList(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With

End Function

Sub DoPasosFx(ByVal CharIndex As Integer)

    With CharList(CharIndex)
        
        If Not UserNavegando Then
            If Not .Muerto And EstaPCarea(CharIndex) And Not esGM(CharIndex) Then
                .Pie = Not .Pie

                Call Audio.PlayEffect(IIf(.Pie, SND_PASOS1, SND_PASOS2), .SoundSource)
            End If
        Else
            ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
            Call Audio.PlayEffect(SND_NAVEGANDO, .SoundSource)
        End If
    End With

End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

        '<EhHeader>
        On Error GoTo MoveCharbyPos_Err

        '</EhHeader>

        Dim X        As Integer

        Dim Y        As Integer

        Dim AddX     As Integer

        Dim AddY     As Integer

        Dim nHeading As E_Heading
    
100     With CharList(CharIndex)
102         X = .Pos.X
104         Y = .Pos.Y
            
            If X > 0 Then
106             If (MapData(X, Y).CharIndex = CharIndex) Then
108                 MapData(X, Y).CharIndex = 0
    
                End If
            End If
            
110         AddX = nX - X
112         AddY = nY - Y
        
114         If Sgn(AddX) = 1 Then
116             nHeading = E_Heading.EAST
118         ElseIf Sgn(AddX) = -1 Then
120             nHeading = E_Heading.WEST
122         ElseIf Sgn(AddY) = -1 Then
124             nHeading = E_Heading.NORTH
126         ElseIf Sgn(AddY) = 1 Then
128             nHeading = E_Heading.SOUTH

            End If
            
            If nHeading = 0 Then Exit Sub
             
130         MapData(nX, nY).CharIndex = CharIndex

132         .Pos.X = nX
134         .Pos.Y = nY
        
136         Call g_Swarm.Move(CharIndex, nX, nY)

138         .MoveOffsetX = -1 * (TilePixelWidth * AddX)
140         .MoveOffsetY = -1 * (TilePixelHeight * AddY)
            
              Call Audio.UpdateSource(.SoundSource, nX, nY)

144         .Heading = nHeading
            
146         .scrollDirectionX = Sgn(AddX)
148         .scrollDirectionY = Sgn(AddY)
            .LastStep = FrameTime
            
            .Idle = False
            
            If Not .Moving Then
        
                'If .Muerto Then
                    '.Body = BodyData(CASPER_BODY)

                'End If
        
                'Start animations
                If .Body.Walk(.Heading).started = 0 Then
                    .Body.Walk(.Heading).started = FrameTime
                    .Arma.WeaponWalk(.Heading).started = FrameTime
                    .Escudo.ShieldWalk(.Heading).started = FrameTime

                    .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                    .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS

                End If
            
                .MovArmaEscudo = False
                .Moving = True

            End If

        End With
    
164     If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)

        '<EhFooter>
        Exit Sub

MoveCharbyPos_Err:
        LogError err.Description & vbCrLf & "in MoveCharbyPos " & "at line " & Erl

        '</EhFooter>
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
        '<EhHeader>
        On Error GoTo MoveScreen_Err
        '</EhHeader>

        '******************************************
        'Starts the screen moving in a direction
        '******************************************
        Dim X  As Integer
        Dim Y  As Integer
        Dim tX As Integer
        Dim tY As Integer
    
        'Figure out which way to move
100     Select Case nHeading

            Case E_Heading.NORTH
102             Y = -1
        
104         Case E_Heading.EAST
106             X = 1
        
108         Case E_Heading.SOUTH
110             Y = 1
        
112         Case E_Heading.WEST
114             X = -1
        End Select
    
    
        'Fill temp pos
116     tX = UserPos.X + X
118     tY = UserPos.Y + Y
    
        'Check to see if its out of bounds
120     If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then

            Exit Sub

        Else
            'Start moving... MainLoop does the rest
122         AddtoUserPos.X = X
124         UserPos.X = tX
126         AddtoUserPos.Y = Y
128         UserPos.Y = tY
130         UserMoving = True
        
132         bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        End If

        
        '<EhFooter>
        Exit Sub

MoveScreen_Err:
        LogError err.Description & vbCrLf & _
           "in MoveScreen " & _
           "at line " & Erl

        '</EhFooter>
End Sub

Function NextOpenChar() As Integer

    '*****************************************************************
    'Finds next open char slot in CharList
    '*****************************************************************
    Dim LoopC As Long

    Dim Dale  As Boolean
    
    LoopC = 1

    Do While CharList(LoopC).Active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(CharList))
    Loop
    
    NextOpenChar = LoopC
End Function

Public Function Graphic_Is_Alpha(ByVal Graphic As Long) As Boolean

    If Graphic >= 60 And Graphic <= 76 Then
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic >= 259 And Graphic <= 269 Then
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic >= 2527 And Graphic <= 2531 Then
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic = 408 Or Graphic = 409 Or Graphic = 2483 Or Graphic = 2497 Or _
       Graphic = 2498 Or Graphic = 2528 Or Graphic = 2519 Or Graphic = 2522 Or Graphic = 25075 Then
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic = 742 Or Graphic = 818 Or Graphic = 888 Or Graphic = 990 Or Graphic = 991 Or Graphic = 1036 Or Graphic = 1037 Or _
       Graphic = 1041 Or Graphic = 1639 Or Graphic = 1640 Or Graphic = 1641 Or Graphic = 1907 Or Graphic = 2522 Or Graphic = 2562 Or _
       Graphic = 2563 Or Graphic = 2564 Or Graphic = 2567 Or Graphic = 2568 Or Graphic = 2569 Or Graphic = 3045 Or Graphic = 3046 Or _
       Graphic = 6041 Or Graphic = 6057 Or Graphic = 6267 Or Graphic = 6270 Or Graphic = 6271 Or Graphic = 6279 Or Graphic = 6280 Or _
       Graphic = 6283 Or Graphic = 6336 Then
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If (Graphic >= 4 And Graphic <= 19) Or _
       (Graphic >= 26 And Graphic <= 34) Or _
       (Graphic >= 441 And Graphic <= 451) Or _
       (Graphic >= 457 And Graphic <= 459) Or _
       (Graphic >= 464 And Graphic <= 499) Or _
       (Graphic >= 858 And Graphic <= 865) Or _
       (Graphic >= 901 And Graphic <= 909) Or _
       (Graphic >= 996 And Graphic <= 1005) Or _
       (Graphic >= 1535 And Graphic <= 1635) Or _
       (Graphic >= 1645 And Graphic <= 1693) Or _
       (Graphic >= 1698 And Graphic <= 1704) Or _
       (Graphic >= 1710 And Graphic <= 1713) Or _
       (Graphic >= 1886 And Graphic <= 1898) Or _
       (Graphic >= 1922 And Graphic <= 1927) Or _
       (Graphic >= 1935 And Graphic <= 1949) Or _
       (Graphic >= 1951 And Graphic <= 1956) Or _
       (Graphic >= 1968 And Graphic <= 1972) Or _
       (Graphic >= 1983 And Graphic <= 1995) Or _
       (Graphic >= 2414 And Graphic <= 2518) Or _
       (Graphic >= 2570 And Graphic <= 2578) Or _
       (Graphic >= 2581 And Graphic <= 2586) Or _
       (Graphic >= 2997 And Graphic <= 3008) Or _
       (Graphic >= 3011 And Graphic <= 3014) Or _
       (Graphic >= 3027 And Graphic <= 3042) Or _
       (Graphic >= 6014 And Graphic <= 6017) Then
        Graphic_Is_Alpha = True
        Exit Function
    End If

    If (Graphic >= 6046 And Graphic <= 6054) Or _
       (Graphic >= 6061 And Graphic <= 6098) Or _
       (Graphic >= 6110 And Graphic <= 6126) Or _
       (Graphic >= 6154 And Graphic <= 6216) Or _
       (Graphic >= 6227 And Graphic <= 6261) Or _
       (Graphic >= 6312 And Graphic <= 6331) Or _
       (Graphic >= 6340 And Graphic <= 6369) Or _
       (Graphic >= 6821 And Graphic <= 6843) Then
        Graphic_Is_Alpha = True
        Exit Function
    
    End If

    If (Graphic >= 1171 And Graphic <= 1177) Or _
       (Graphic >= 6052 And Graphic <= 6096) Or _
       (Graphic >= 6892 And Graphic <= 6896) Or _
       (Graphic >= 1213 And Graphic <= 1218) Or _
       (Graphic >= 1923 And Graphic <= 1927) Or _
       (Graphic >= 4252 And Graphic <= 4269) Then
        Graphic_Is_Alpha = True
        Exit Function
    
    End If
    
    If Graphic = 1168 Or Graphic = 1179 Or Graphic = 1181 Or Graphic = 1183 Or _
       Graphic = 1184 Or Graphic = 1086 Or Graphic = 1087 Or Graphic = 4958 Or Graphic = 4957 Or _
       Graphic = 1193 Or Graphic = 1194 Or Graphic = 1195 Or _
       Graphic = 1198 Or Graphic = 1199 Or Graphic = 1200 Or _
       Graphic = 6052 Or Graphic = 6063 Or Graphic = 6064 Or _
       Graphic = 6065 Or Graphic = 6066 Or Graphic = 6067 Or _
       Graphic = 6069 Or Graphic = 6071 Or Graphic = 6072 Or Graphic = 6073 Or Graphic = 6074 Or _
       Graphic = 6075 Or Graphic = 6078 Or Graphic = 6079 Or Graphic = 6081 Or Graphic = 6082 Or _
       Graphic = 6083 Or Graphic = 6084 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    
    If Graphic = 6085 Or Graphic = 6086 Or Graphic = 6087 Or Graphic = 6088 Or _
       Graphic = 6103 Or Graphic = 6117 Or Graphic = 6154 Or Graphic = 4912 Or _
       Graphic = 6155 Or Graphic = 6156 Or Graphic = 6157 Or _
       Graphic = 6518 Or Graphic = 6160 Or Graphic = 6161 Or _
       Graphic = 6162 Or Graphic = 6163 Or Graphic = 6164 Or _
       Graphic = 6165 Or Graphic = 6166 Or Graphic = 6167 Or Graphic = 5043 Or _
       Graphic = 6168 Or Graphic = 6174 Or Graphic = 6175 Or Graphic = 6176 Or _
       Graphic = 6177 Or Graphic = 6178 Or Graphic = 6193 Or Graphic = 6194 Or Graphic = 6197 Or _
       Graphic = 6227 Or Graphic = 6228 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic = 6229 Or Graphic = 888 Or Graphic = 997 Or Graphic = 1005 Or _
       Graphic = 996 Or Graphic = 1212 Or Graphic = 1602 Or _
       Graphic = 1605 Or Graphic = 1922 Or Graphic = 1944 Or _
       Graphic = 1951 Or Graphic = 1952 Or Graphic = 1949 Or _
       Graphic = 2569 Or Graphic = 3689 Or Graphic = 3690 Or _
       Graphic = 4193 Or Graphic = 4251 Or Graphic = 4282 Or _
       Graphic = 4283 Or Graphic = 6169 Or Graphic = 4282 Or Graphic = 4283 Or _
       Graphic = 4284 Or Graphic = 4360 Or Graphic = 4374 Or Graphic = 4376 Or Graphic = 4440 Or _
       Graphic = 4753 Or Graphic = 4754 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic = 4759 Or Graphic = 4760 Or Graphic = 4761 Or Graphic = 4801 Or _
       Graphic = 4802 Or Graphic = 4718 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If Graphic = 1910 Or Graphic = 1928 Or Graphic = 1929 Or Graphic = 1930 Or _
       Graphic = 1931 Or Graphic = 3017 Or Graphic = 3019 Or _
       Graphic = 3024 Or Graphic = 3495 Or Graphic = 3669 Or _
       Graphic = 3722 Or Graphic = 4234 Or Graphic = 4235 Or _
       Graphic = 4275 Or Graphic = 4286 Or Graphic = 4287 Or _
       Graphic = 4734 Or Graphic = 4735 Or Graphic = 4736 Or _
       Graphic = 4737 Or Graphic = 4969 Or Graphic = 4971 Or Graphic = 4972 Or _
       Graphic = 4973 Or Graphic = 4976 Or Graphic = 5185 Or Graphic = 5186 Or Graphic = 5207 Or _
       Graphic = 4753 Or Graphic = 4754 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    
    If Graphic = 5271 Or Graphic = 5272 Or Graphic = 5292 Or Graphic = 5323 Or _
       Graphic = 5324 Or Graphic = 5325 Or Graphic = 3019 Or _
       Graphic = 3024 Or Graphic = 3495 Or Graphic = 3669 Or _
       Graphic = 3722 Or Graphic = 4234 Or Graphic = 4235 Or _
       Graphic = 4275 Or Graphic = 4286 Or Graphic = 4287 Or _
       Graphic = 4734 Or Graphic = 4735 Or Graphic = 4736 Or _
       Graphic = 4737 Or Graphic = 4969 Or Graphic = 4971 Or Graphic = 4972 Or _
       Graphic = 4973 Or Graphic = 4976 Or Graphic = 4977 Or Graphic = 5185 Or Graphic = 5186 Or Graphic = 5207 Or _
       Graphic = 4753 Or Graphic = 4754 Or Graphic = 5271 Or Graphic = 5272 Or Graphic = 5292 Or _
       Graphic = 6097 Or Graphic = 6281 Or Graphic = 6278 Or Graphic = 1168 Or Graphic = 1174 Or Graphic = 1175 Or _
       Graphic = 1185 Or Graphic = 1186 Or Graphic = 1187 Or Graphic = 1198 Or Graphic = 2001 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If (Graphic >= 5323 And Graphic <= 5339) Or _
        (Graphic >= 6128 And Graphic <= 6153) Then
        Graphic_Is_Alpha = True
        Exit Function
    
    End If
    
    
    
     If Graphic = 212 Or Graphic = 1213 Or Graphic = 1214 Or Graphic = 1215 Or Graphic = 1216 Or Graphic = 1217 Or _
       Graphic = 1714 Or Graphic = 1715 Or Graphic = 1716 Or Graphic = 1717 Or Graphic = 1718 Or Graphic = 1719 Or _
       Graphic = 1730 Or Graphic = 1731 Or Graphic = 1732 Or Graphic = 1733 Or Graphic = 1735 Or Graphic = 1736 Or _
       Graphic = 1737 Or Graphic = 1738 Or Graphic = 1739 Or _
       Graphic = 1740 Or Graphic = 1741 Or Graphic = 1745 Or Graphic = 1746 Or Graphic = 1748 Or Graphic = 1796 Or _
       Graphic = 1797 Or Graphic = 1798 Or Graphic = 1805 Or Graphic = 6964 Or _
       Graphic = 1806 Or Graphic = 1807 Or Graphic = 1808 Or Graphic = 1809 Or _
       Graphic = 1810 Or Graphic = 1811 Or Graphic = 1812 Or Graphic = 1813 Or Graphic = 1814 Or Graphic = 1815 Or _
       Graphic = 1816 Or Graphic = 1817 Or Graphic = 1818 Or Graphic = 1819 Or Graphic = 1820 Or _
       Graphic = 1821 Or Graphic = 1826 Or Graphic = 1827 Or Graphic = 1828 Or Graphic = 1829 Or Graphic = 1830 Or _
       Graphic = 1841 Or Graphic = 1842 Or Graphic = 1843 Or Graphic = 1850 Or Graphic = 1851 Or Graphic = 1852 Or _
       Graphic = 1853 Or Graphic = 1854 Or Graphic = 1855 Or Graphic = 1856 Or Graphic = 1857 Or Graphic = 1865 Or _
       Graphic = 1866 Or Graphic = 1867 Or Graphic = 1868 Or Graphic = 1912 Or Graphic = 1917 Or Graphic = 1918 Or _
       Graphic = 1979 Or Graphic = 1980 Or Graphic = 3040 Or Graphic = 4598 Or Graphic = 4599 Or Graphic = 4600 Or _
       Graphic = 4838 Or Graphic = 4923 Or Graphic = 4927 Or Graphic = 5009 Or Graphic = 6145 Or Graphic = 6963 Or _
       Graphic = 1835 Or Graphic = 1836 Or Graphic = 1838 Or Graphic = 1839 Or Graphic = 1840 Then
        
        Graphic_Is_Alpha = True
        Exit Function
    End If
    
    If (Graphic >= 4602 And Graphic <= 4615) Or _
        (Graphic >= 4882 And Graphic <= 4914) Or _
        (Graphic >= 4940 And Graphic <= 4949) Or _
        (Graphic >= 4957 And Graphic <= 4969) Or _
        (Graphic >= 5036 And Graphic <= 5053) Or _
        (Graphic >= 4857 And Graphic <= 4871) Then
        Graphic_Is_Alpha = True
        Exit Function
    
    End If
    

End Function

Public Function GrhData_Type(ByVal GrhIndex As Long) As eGrhType
        
        Dim Temp As eGrhType


            ' Arboles Transparentes
            Select Case GrhIndex
                Case 7000, 7001, 7002, 647, 70
                    Temp = eGrhType.eArbol
                
                Case Else
                    Temp = eGrhType.eNone
            End Select
    
    GrhData_Type = Temp
        
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Public Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim grh         As Long

    Dim Frame       As Long

    Dim handle      As Integer

    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , MaxGrh
    
    'Resize arrays
    ReDim GrhData(0 To MaxGrh) As GrhData
    ReDim GrhDataDefault(0 To MaxGrh) As GrhData
    
    Dim Mult
    
    #If ModoBig > 0 Then
        Mult = 2
    #Else
        Mult = 1
    #End If

    While Not EOF(handle)

        Get handle, , grh
        
        If grh <> 0 Then
            
            With GrhData(grh)
                
                ' Set de Type Grh
                .GrhType = GrhData_Type(grh)
                GrhDataDefault(grh).GrhType = .GrhType
                
                'Get number of frames
                Get handle, , .NumFrames
                GrhDataDefault(grh).NumFrames = .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                ReDim GrhDataDefault(grh).Frames(1 To .NumFrames)

                If .NumFrames > 1 Then

                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        GrhDataDefault(grh).Frames(Frame) = .Frames(Frame)
                        
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > MaxGrh Then
                            GoTo ErrorHandler

                        End If

                    Next Frame
                    
                    Get handle, , .Speed
                    GrhDataDefault(grh).Speed = .Speed
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    
                    'Compute width and height
                    GrhDataDefault(grh).pixelHeight = GrhData(.Frames(1)).pixelHeight
                    GrhData(grh).pixelHeight = GrhDataDefault(grh).pixelHeight * Mult

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    GrhDataDefault(grh).pixelWidth = GrhData(.Frames(1)).pixelWidth
                    GrhData(grh).pixelWidth = GrhDataDefault(grh).pixelWidth * Mult
                    
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    GrhDataDefault(grh).TileWidth = GrhData(.Frames(1)).TileWidth
                    GrhData(grh).TileWidth = GrhDataDefault(grh).TileWidth * Mult

                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    GrhDataDefault(grh).TileHeight = GrhData(.Frames(1)).TileHeight
                    GrhData(grh).TileHeight = GrhDataDefault(grh).TileHeight * Mult

                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    GrhDataDefault(grh).FileNum = .FileNum
                    If MAX_TEXTURE < .FileNum Then
                        MAX_TEXTURE = .FileNum

                    End If
                    
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    .Alpha = Graphic_Is_Alpha(.FileNum)
                    GrhDataDefault(grh).Alpha = .Alpha
                    Get handle, , .sX
                    
                    GrhDataDefault(grh).sX = .sX
                    .sX = GrhDataDefault(grh).sX * Mult

                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    GrhDataDefault(grh).sY = .sY
                    .sY = GrhDataDefault(grh).sY * Mult

                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    GrhDataDefault(grh).pixelWidth = .pixelWidth
                    .pixelWidth = GrhDataDefault(grh).pixelWidth * Mult

                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    GrhDataDefault(grh).pixelHeight = .pixelHeight
                    .pixelHeight = GrhDataDefault(grh).pixelHeight * Mult
                    
                    Get handle, , .src.x1
                    Get handle, , .src.y1
                    Get handle, , .src.X2
                    Get handle, , .src.Y2
                    
                    .src.x1 = .src.x1
                    .src.y1 = .src.y1
                    .src.X2 = .src.X2
                    .src.Y2 = .src.Y2
                    GrhDataDefault(grh).src = .src
                    
                    'Compute width and height
                    .TileWidth = (.pixelWidth / (TilePixelWidth))
                    .TileHeight = (.pixelHeight / (TilePixelHeight))
                    
                    .Frames(1) = grh
                    
                    GrhDataDefault(grh).TileWidth = .TileWidth
                    GrhDataDefault(grh).TileHeight = .TileHeight
                    GrhDataDefault(grh).Frames(1) = grh
                End If

            End With

        End If

    Wend
    
    Close handle

    ReDim g_Textures(1 To MAX_TEXTURE) As Integer
    ReDim g_Textures_Gui(1 To MAX_TEXTURE_GUI) As Integer
    ReDim g_Textures_MiniMapa(1 To 200) As Integer
    ReDim g_Textures_Avatars(0 To MAX_TEXTURE_AVATARS) As Integer
    LoadGrhData = True

    Exit Function

ErrorHandler:
    LoadGrhData = False

End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then

        Exit Function

    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then

        Exit Function

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then

        Exit Function

    End If
    
    'If UserMontando = HayAgua(X, Y) Then
    'Exit Function
    'End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 01/08/2009
    'Checks to see if a tile position is legal, including if there is a casper in the tile
    '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
    '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
    '*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then

        Exit Function

    End If
    
    CharIndex = MapData(X, Y).CharIndex

    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then

            Exit Function

        End If
        
        With CharList(CharIndex)

            If (.iBody <> iCuerpoMuerto And .iBody <> iCuerpoMuerto_Legion And .iBody <> FRAGATA_FANTASMAL) Or .NpcIndex > 0 Then
                    
                Exit Function

            End If
            
            
            ' Caspers no pueden ser ser pasados si estan "meditando"
            If .iBody = iCuerpoMuerto Or .iBody = FRAGATA_FANTASMAL Or .iBody <> iCuerpoMuerto_Legion Then
                
                If .FxIndex > 0 Then
                    Exit Function
                End If
            End If
            
            ' Si no es casper, no puede pasar
            'If .iBody <> CUERPO_FANTASMAL And .iBody <> FRAGATA_FANTASMAL Then

            'Exit Function

            '   Else

            ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
            If HayAgua(UserPos.X, UserPos.Y) Then
                If Not HayAgua(X, Y) Then Exit Function
            Else

                ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                If HayAgua(X, Y) Then Exit Function

            End If
                
            ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
            If CharList(UserCharIndex).Priv > 0 And CharList(UserCharIndex).Priv < 5 Then
                If CharList(UserCharIndex).Invisible = True Then Exit Function

            End If

            'End If

        End With

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then

        Exit Function

    End If
    
    'If UserMontando = HayAgua(X, Y) Then
    '    Exit Function
    'End If
    
    MoveToLegalPos = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub RenderScreen_FullScreen()
        
    Dim X    As Long, Y As Long

    Dim Mult As Byte
    
    ' fps
    Mult = 2
    

    Draw_Text f_Tahoma, 12, X * Mult, Y * Mult, To_Depth(9), 0, ARGB(255, 255, 255, 100), FONT_ALIGNMENT_CENTER, UserFps, False
    
    
    Dim Width As Long
    Dim Height As Long
    
    Width = FrmMain.MainViewPic.ScaleWidth
    Height = FrmMain.MainViewPic.ScaleHeight
    
    
    ' Mapa
    X = Width - 320
    Y = 30
    Call Draw_Texture_Graphic_Gui(139, X, Y, To_Depth(8), 293, 33, 0, 0, 293, 33, -1, 0, eTechnique.t_Alpha)
    Call Draw_Text(f_Tahoma, 10, X + 137, Y + 17, To_Depth(9), 0, -1, FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_MIDDLE, MiniMap(UserMap).Name, False)
    
    ' Usuarios Online
    X = Width - 123
    Y = 80
    Call Draw_Texture_Graphic_Gui(140, X, Y, To_Depth(8), 97, 97, 0, 0, 97, 97, -1, 0, eTechnique.t_Alpha)
    Call Draw_Text(f_Tahoma, 15, X + 48, Y + 44, To_Depth(9), 0, ARGB(45, 255, 70, 100), FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_MIDDLE, UsuariosOnline, False)
    Call Draw_Text(f_Tahoma, 10, X + 48, Y + 66, To_Depth(9), 0, ARGB(45, 150, 70, 100), FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_MIDDLE, "Ons", False)
    
    X = 750 * Mult
    Y = 50 * Mult
    ' Draw_Text f_Verdana, 15, X, Y, To_Depth(9), 0, ARGB(45, 255, 70, 100), FONT_ALIGNMENT_CENTER, UsuariosOnline, False
  '  Draw_Text f_Medieval, 30, X, Y, To_Depth(9), 0, ARGB(45, 255, 70, 255), FONT_ALIGNMENT_RIGHT, "Online: " & UsuariosOnline, True
    'Call Draw_Texture_Graphic_Gui(8, X - 360, Y - 42, To_Depth(8), 415, 60, 0, 0, 415, 60, ARGB(200, 6, 6, 120), 0, eTechnique.t_Alpha)
    
    ' comando !LIVE
    'X = 10 * Mult
    'Y = 500 * Mult
    ' Draw_Text f_Verdana, 15, X, Y, To_Depth(9), 0, ARGB(45, 255, 70, 255), FONT_ALIGNMENT_LEFT, "Tipea '/LIVE' dentro del juego para aparecer en el stream.", True
    
End Sub

' # Sistema de mini mapa con movimiento
Sub RenderScreen_MiniMap()

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
        
    Dim Mult       As Byte
        
    #If ModoBig > 0 Then
        Mult = 2
    #Else
        Mult = 1
    #End If
    
        
     Call wGL_Graphic.Use_Device(g_Captions(eCaption.cMiniMapa))
     Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmMain.MiniMapa.ScaleWidth * 4, FrmMain.MiniMapa.ScaleHeight * 4)
     Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
     
  If UserMoving Then
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame * CharList(UserCharIndex).Speeding

            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False

            End If

        End If

        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * CharList(UserCharIndex).Speeding

            If Abs(OffsetCounterY) >= Abs(TilePixelWidth * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = False

            End If

        End If

    End If
    
    'Figure out Ends and Starts of screen
    ScreenMinY = (UserPos.Y - AddtoUserPos.Y) - HalfWindowTileHeight
    ScreenMaxY = (UserPos.Y - AddtoUserPos.Y) + HalfWindowTileHeight
    ScreenMinX = (UserPos.X - AddtoUserPos.X) - HalfWindowTileWidth
    ScreenMaxX = (UserPos.X - AddtoUserPos.X) + HalfWindowTileWidth
    
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

    For Y = YMinMapSize To YMaxMapSize
        DrawableY = (Y) * TilePixelWidth
    
        For X = XMinMapSize To XMaxMapSize
            DrawableX = (X) * TilePixelHeight

            With MapData(X, Y)

                If (.Graphic(1).GrhIndex <> 0) Then
                    Call Draw_Grh(.Graphic(1), DrawableX, DrawableY, To_Depth(1, X, Y), 0, 1)

                End If

            End With

        Next X
    Next Y

    Dim Results() As wGL_Swarm_Result

    Call g_Swarm.Query(MinX, MinY, MaxX, MaxY, Results)
        
    Call SetAlpha(RoofAlpha, bTecho, ROOF_ALPHA_MIN, ROOF_ALPHA_MAX, ROOF_ALPHA_SPEED)
    
    For Drawable = 0 To UBound(Results)

        With Results(Drawable)
    
            DrawableX = (.X - ScreenMinX) * TilePixelWidth + OffsetCounterX
            DrawableY = (.Y - ScreenMinY) * TilePixelHeight + OffsetCounterY
                    
            Select Case (.Layer)

                Case 1
                   ' Call Draw_Grh(MapData(.X, .Y).Graphic(2), DrawableX, DrawableY, To_Depth(2, .X, .Y), 1, 1, , , , eTechnique.t_Alpha)

                Case 2

                    If GrhData(MapData(.X, .Y).Graphic(3).GrhIndex).GrhType = eGrhType.eArbol Then
                        If CharList(UserCharIndex).Pos.X >= .X - 3 And CharList(UserCharIndex).Pos.X <= .Y + 3 Then
                          '  Call Draw_Grh(MapData(.X, .Y).Graphic(3), DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , ARGB(255, 255, 255, 80), , eTechnique.t_Alpha)
                        Else
                           ' Call Draw_Grh(MapData(.X, .Y).Graphic(3), DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , , , eTechnique.t_Alpha)

                        End If

                    Else
                       ' Call Draw_Grh(MapData(.X, .Y).Graphic(3), DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , , , eTechnique.t_Alpha)

                    End If

                Case 3
                        
                    'Call Draw_Grh(MapData(.X, .Y).Graphic(4), DrawableX, DrawableY, To_Depth(4, .X, .Y), 1, 1, , ARGB(255, 255, 255, RoofAlpha), , eTechnique.t_Alpha)
                        
                Case 4

                    If GrhData(MapData(.X, .Y).ObjGrh.GrhIndex).GrhType = eGrhType.eArbol Then
                        If CharList(UserCharIndex).Pos.X >= .X - 3 And CharList(UserCharIndex).Pos.X <= .Y + 3 Then
                         '   Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 1), 1, 1, , ARGB(255, 255, 255, 120), , eTechnique.t_Alpha)
                        Else
                            'Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 1), 1, 1, , , , eTechnique.t_Alpha)

                        End If

                    Else
                       ' Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , , , eTechnique.t_Alpha)
                    End If

                Case 5

                    If UserCharIndex = MapData(.X, .Y).CharIndex Then
                        Call CharRender(MapData(.X, .Y).CharIndex, DrawableX, DrawableY, Mult)
                        CharList(MapData(.X, .Y).CharIndex).NowPosX = DrawableX
                        CharList(MapData(.X, .Y).CharIndex).NowPosY = DrawableY

                    End If

            End Select

        End With

    Next Drawable


    Call wGL_Graphic_Renderer.Flush
End Sub

Sub RenderScreen()

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
        
    Dim Mult       As Byte
        
    #If ModoBig > 0 Then
        Mult = 2
    #Else
        Mult = 1
    #End If

    
  If UserMoving Then
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame * CharList(UserCharIndex).Speeding

            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False

            End If

        End If

        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * CharList(UserCharIndex).Speeding

            If Abs(OffsetCounterY) >= Abs(TilePixelWidth * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = False

            End If

        End If

    End If
    
    'Figure out Ends and Starts of screen
    ScreenMinY = (UserPos.Y - AddtoUserPos.Y) - HalfWindowTileHeight
    ScreenMaxY = (UserPos.Y - AddtoUserPos.Y) + HalfWindowTileHeight
    ScreenMinX = (UserPos.X - AddtoUserPos.X) - HalfWindowTileWidth
    ScreenMaxX = (UserPos.X - AddtoUserPos.X) + HalfWindowTileWidth
    
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

    For Y = MinY To MaxY
        DrawableY = (Y - ScreenMinY) * TilePixelWidth + OffsetCounterY
    
        For X = MinX To MaxX
            DrawableX = (X - ScreenMinX) * TilePixelHeight + OffsetCounterX

            With MapData(X, Y)

                If (.Graphic(1).GrhIndex <> 0) Then
                    Call Draw_Grh(.Graphic(1), DrawableX, DrawableY, To_Depth(1, X, Y), 0, 1)

                End If
                    
                If (.Damage.DamageType <> 0) Then
                    mDamages.DrawDamage X, Y, DrawableX + 17, DrawableY - 2

                End If

            End With

        Next X
    Next Y

    Dim Results() As wGL_Swarm_Result

    Call g_Swarm.Query(MinX, MinY, MaxX, MaxY, Results)
        
    Call SetAlpha(RoofAlpha, bTecho, ROOF_ALPHA_MIN, ROOF_ALPHA_MAX, ROOF_ALPHA_SPEED)
    
    For Drawable = 0 To UBound(Results)

        With Results(Drawable)
    
            DrawableX = (.X - ScreenMinX) * TilePixelWidth + OffsetCounterX
            DrawableY = (.Y - ScreenMinY) * TilePixelHeight + OffsetCounterY
                    
            Select Case (.Layer)

                Case 1
                    Call Draw_Grh(MapData(.X, .Y).Graphic(2), DrawableX, DrawableY, To_Depth(2, .X, .Y), 1, 1, , , , eTechnique.t_Alpha)

                Case 2

                    If GrhData(MapData(.X, .Y).Graphic(3).GrhIndex).GrhType = eGrhType.eArbol Then
                        If CharList(UserCharIndex).Pos.X >= .X - 3 And CharList(UserCharIndex).Pos.X <= .Y + 3 Then
                            Call Draw_Grh(MapData(.X, .Y).Graphic(3), DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , ARGB(255, 255, 255, 80), , eTechnique.t_Alpha)
                        Else
                            Call Draw_Grh(MapData(.X, .Y).Graphic(3), DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , , , eTechnique.t_Alpha)

                        End If

                    Else
                        Call Draw_Grh(MapData(.X, .Y).Graphic(3), DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , , , eTechnique.t_Alpha)

                    End If

                Case 3
                        
                    Call Draw_Grh(MapData(.X, .Y).Graphic(4), DrawableX, DrawableY, To_Depth(4, .X, .Y), 1, 1, , ARGB(255, 255, 255, RoofAlpha), , eTechnique.t_Alpha)
                        
                Case 4

                    ' @ Luz sobre los objetos
                    If MapData(.X, .Y).OBJInfo.ObjIndex > 0 Then
                        If ObjData(MapData(.X, .Y).OBJInfo.ObjIndex).Color <> 0 Then
                            Call Draw_Texture_Graphic_Gui(4, DrawableX, DrawableY, To_Depth(3, .X, .Y, 1), 32 * Mult, 32 * Mult, 0, 0, 32 * Mult, 32 * Mult, ObjData(MapData(.X, .Y).OBJInfo.ObjIndex).Color, 0, eTechnique.t_Alpha)

                        End If

                    End If

                    If GrhData(MapData(.X, .Y).ObjGrh.GrhIndex).GrhType = eGrhType.eArbol Then
                        If CharList(UserCharIndex).Pos.X >= .X - 3 And CharList(UserCharIndex).Pos.X <= .Y + 3 Then
                            Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 1), 1, 1, , ARGB(255, 255, 255, 120), , eTechnique.t_Alpha)
                        Else
                            Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 1), 1, 1, , , , eTechnique.t_Alpha)

                        End If

                    Else
                        Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 2), 1, 1, , , , eTechnique.t_Alpha)

                    End If
                        
                    'Call Draw_Grh(MapData(.X, .Y).ObjGrh, DrawableX, DrawableY, To_Depth(3, .X, .Y, 1), 1, 1, , , , eTechnique.t_Alpha)

                Case 5

                    ' Debug.Assert MapData(X, Y).CharIndex <> 0
                    If MapData(.X, .Y).CharIndex <> 0 Then
                        Call CharRender(MapData(.X, .Y).CharIndex, DrawableX, DrawableY, Mult)
                        CharList(MapData(.X, .Y).CharIndex).NowPosX = DrawableX
                        CharList(MapData(.X, .Y).CharIndex).NowPosY = DrawableY

                    End If

                    'Call Effect_Render_All
              
                Case 6

                    ' Fxs sobre el terreneitor
                    Call Draw_Grh(MapData(.X, .Y).fX, DrawableX + FxData(MapData(.X, .Y).FxIndex).OffsetX, DrawableY + FxData(MapData(.X, .Y).FxIndex).OffsetY, To_Depth(4), 1, 1, 1, ARGB(255, 255, 255, ClientSetup.bAlpha), , eTechnique.t_Alpha)
                                    
                    'Check if animation is over
                    If MapData(.X, .Y).fX.started = 0 Then
                        MapData(.X, .Y).FxIndex = 0
    
                        Call g_Swarm.Remove(6, -1, .X, .Y, 2, 2)

                    End If
                        
                Case 7

                    Dim Value As Long

                    Value = GrhData(BAR_BACKGROUND).pixelWidth - (((MapData(.X, .Y).BarMin / 100) / (MapData(.X, .Y).BarMax / 100)) * GrhData(BAR_BACKGROUND).pixelWidth)
                            
                    Draw_Texture BAR_BORDER, DrawableX, DrawableY - (35 * Mult), To_Depth(7, .X, .Y), GrhData(BAR_BORDER).pixelWidth, GrhData(BAR_BORDER).pixelHeight, -1, 0, 0
                    Draw_Texture BAR_BACKGROUND, DrawableX + (2 * Mult), DrawableY - (33 * Mult), To_Depth(7, .X, .Y, 3), Value, GrhData(BAR_BACKGROUND).pixelHeight, -1, 0, 0
                    
            End Select

        End With

    Next Drawable
    
    
    ' Consola Flotante
    Call RenderText_Console
        
    Dim X_Temp As Integer

    Dim Y_Temp As Integer

    X_Temp = -70


    If CountDownTime > 0 Then
        If CountDownTime_Fight Then
            CountDownTime = CountDownTime - 1
            Draw_Text f_Morpheus, 150, 270 * Mult, 120 * Mult, To_Depth(6), 0, ARGB(24, 240, 10, 1 + CountDownTime), FONT_ALIGNMENT_CENTER, "¡YA!", False
        Else
            Draw_Text f_Morpheus, 200, 270 * Mult, 120 * Mult, To_Depth(6), 0, ARGB(255, 255, 255, 200), FONT_ALIGNMENT_CENTER, CountDownTime, False

        End If

    End If

    If Map_TimeRender > 0 Then
        If Map_TimeRender = 2000 Then
            Draw_Text f_Medieval, 50, 300 * Mult, 30 * Mult, To_Depth(6), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_TOP, UserMapName, False
        Else
            Map_TimeRender = Map_TimeRender - 1
            Draw_Text f_Medieval, 50, 300 * Mult, 30 * Mult, To_Depth(6), 0, ARGB(255, 255, 255, Map_TimeRender), FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_TOP, UserMapName, False

        End If
            
    End If
    
    Dim TemporalX As Integer
        
    Y_Temp = -64
        
    TemporalX = 550 + X_Temp
        
    Call Anuncio_Update_Render
        
        
        
    #If FullScreen = 1 Then
        
        'Call RenderScreen_FullScreen
    #End If
    
End Sub

Public Sub SetAlpha(ByRef Alpha As Single, ByVal IsActive As Boolean, ByVal min As Byte, ByVal max As Byte, ByVal Speed As Single)
        Dim Value As Single
        
        If (IsActive) Then
            Value = Alpha - Speed * timerElapsedTime
            If (Value < min) Then Alpha = min Else Alpha = Value
        Else
            Value = Alpha + Speed * timerElapsedTime
            If (Value > max) Then Alpha = max Else Alpha = Value
        End If
End Sub

'
' Dibujamos un Grh
'
Public Sub Draw_Grh(ByRef grh As grh, _
                    ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Z As Single, _
                    ByVal Center As Byte, _
                    Optional ByVal Animate As Byte = 0, _
                    Optional ByVal killAtEnd As Byte = 0, _
                    Optional ByVal Colour As Long = -1, _
                    Optional ByVal Rotation As Single = 0, _
                    Optional ByVal Technique As Integer = eTechnique.t_Default, _
                    Optional ByVal pixelWidth As Integer = 0, _
                    Optional ByVal pixelHeight As Integer = 0, _
                    Optional ByVal FxSlot As Byte = 0, _
                    Optional NotBig As Boolean = False, _
                    Optional ByVal SoloBig As Boolean = False)
        '<EhHeader>
        On Error GoTo Draw_Grh_Err
        '</EhHeader>
    
        Dim CurrentGrhIndex As Long

100     If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
        Dim CurrentFrame As Integer

102     CurrentFrame = 1

104     If Animate Then
106         If grh.started > 0 Then

                Dim ElapsedFrames As Long

108             ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.Speed)

110             If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
112                 CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

                Else
114                 grh.started = 0

                End If

            End If

        End If

        'Figure out what frame to draw (always 1 if not animated)
116     CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)
    
        Dim W As Long

        Dim H As Long

118     W = TilePixelWidth
120     H = TilePixelHeight
            
122     With GrhData(CurrentGrhIndex)
            #If ModoBig > 0 Then

124             If NotBig Then
126                 W = TilePixelWidth / 2
128                 H = TilePixelHeight / 2
                
                End If
            
            #Else

130             If SoloBig = True Then

                    ' W = TilePixelWidth * 2
                    ' H = TilePixelHeight * 2
                End If

            #End If
        
132         If Center Then
134             If .TileWidth <> 1 Then
136                 X = X - Int(.TileWidth * W * 0.5) + W * 0.5

                End If
                
138             If .TileHeight <> 1 Then
140                 Y = Y - Int(.TileHeight * H) + H

                End If

            End If

142         If pixelWidth <> 0 And pixelHeight <> 0 Then
144             Draw_Texture CurrentGrhIndex, X, Y, Z, pixelWidth, pixelHeight, Colour, Rotation, Technique
            Else
146             Draw_Texture CurrentGrhIndex, X, Y, Z, .pixelWidth, .pixelHeight, Colour, Rotation, Technique

            End If

        End With



        '<EhFooter>
        Exit Sub

Draw_Grh_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Mod_TileEngine.Draw_Grh " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
'
' Dibujamos un Grh en x1
'
Public Sub Draw_Grh_Default(ByRef grh As grh, _
                    ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Z As Single, _
                    ByVal Center As Byte, _
                    Optional ByVal Animate As Byte = 0, _
                    Optional ByVal killAtEnd As Byte = 0, _
                    Optional ByVal Colour As Long = -1, _
                    Optional ByVal Rotation As Single = 0, _
                    Optional ByVal Technique As Integer = eTechnique.t_Default, _
                    Optional ByVal pixelWidth As Integer = 0, _
                    Optional ByVal pixelHeight As Integer = 0, _
                    Optional ByVal FxSlot As Byte = 0)
        '<EhHeader>
        On Error GoTo Draw_Grh_Default_Err
        '</EhHeader>
    
        Dim CurrentGrhIndex As Long

100     If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
        Dim CurrentFrame As Integer

102     CurrentFrame = 1

104     If Animate Then
106         If grh.started > 0 Then

                Dim ElapsedFrames As Long

108             ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.Speed)

110             If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhDataDefault(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
112                 CurrentFrame = ElapsedFrames Mod GrhDataDefault(grh.GrhIndex).NumFrames + 1

                Else
114                 grh.started = 0

                End If

            End If

        End If

        'Figure out what frame to draw (always 1 if not animated)
116     CurrentGrhIndex = GrhDataDefault(grh.GrhIndex).Frames(CurrentFrame)
    
        Dim W As Long

        Dim H As Long

118     W = TilePixelWidth
120     H = TilePixelHeight
            
122     With GrhDataDefault(CurrentGrhIndex)
        
132         If Center Then
134             If .TileWidth <> 1 Then
136                 X = X - Int(.TileWidth * W * 0.5) + W * 0.5

                End If
                
138             If .TileHeight <> 1 Then
140                 Y = Y - Int(.TileHeight * H) + H

                End If

            End If

142         If pixelWidth <> 0 And pixelHeight <> 0 Then
144             Draw_Texture_Default CurrentGrhIndex, X, Y, Z, pixelWidth, pixelHeight, Colour, Rotation, Technique
            Else
146             Draw_Texture_Default CurrentGrhIndex, X, Y, Z, .pixelWidth, .pixelHeight, Colour, Rotation, Technique

            End If

        End With



        '<EhFooter>
        Exit Sub

Draw_Grh_Default_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Mod_TileEngine.Draw_Grh_Default " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Public Sub Draw_GrhFX(ByVal CharIndex As Integer, _
                      ByVal FxSlot As Byte, _
                      ByRef grh As grh, _
                      ByVal X As Integer, _
                      ByVal Y As Integer, _
                      ByVal Z As Single, _
                      ByVal Center As Byte, _
                      Optional ByVal Animate As Byte = 0, _
                      Optional ByVal killAtEnd As Byte = 0, _
                      Optional ByVal Colour As Long = -1, _
                      Optional ByVal Rotation As Single = 0, _
                      Optional ByVal Technique As Integer = eTechnique.t_Default, _
                      Optional ByVal pixelWidth As Integer = 0, _
                      Optional ByVal pixelHeight As Integer = 0, _
                      Optional NotBig As Boolean = False, _
                      Optional ByVal SoloBig As Boolean = False)
    
    Dim CurrentGrhIndex As Long

    On Error GoTo error

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer

    CurrentFrame = 1

    If Animate Then
        If grh.started > 0 Then

            Dim ElapsedFrames As Long

            ElapsedFrames = Fix((FrameTime - grh.started) / grh.Speed)

            If grh.AnimacionContador > 0 Then
                grh.AnimacionContador = grh.AnimacionContador - ElapsedFrames

            End If
            
            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0

            End If

        End If

    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)
    
    If grh.AnimacionContador < grh.CantAnim * 0.1 Then
        grh.Alpha = grh.Alpha - 1
    
        If grh.Alpha = 0 And CharIndex > 0 Then
            If FxSlot <> 0 Then
                CharList(CharIndex).FxList(FxSlot).started = 0
            Else
                CharList(CharIndex).fX.started = 0
            End If
            
            Exit Sub

        End If

    End If
    
    If grh.AnimacionContador > grh.CantAnim * 0.6 Then
       If grh.Alpha < 220 Then
            grh.Alpha = grh.Alpha + 1

        End If

    End If
    
    ' Obtener el valor actual del color long
    Dim colorValue As Long

    colorValue = Colour 'ARGB(255, 255, 255, grh.Alpha) ' ejemplo de un color naranja con una transparencia del 50%

    Dim W As Long

    Dim H As Long

    W = TilePixelWidth
    H = TilePixelHeight
            
    With GrhData(CurrentGrhIndex)
        #If ModoBig > 0 Then

            If NotBig Then
                W = TilePixelWidth / 2
                H = TilePixelHeight / 2
                
            End If
            
        #Else

            If SoloBig = True Then

                ' W = TilePixelWidth * 2
                ' H = TilePixelHeight * 2
            End If

        #End If
        
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * W * 0.5) + W * 0.5

            End If
                
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * H) + H

            End If

        End If

        If pixelWidth <> 0 And pixelHeight <> 0 Then
            Draw_Texture CurrentGrhIndex, X, Y, Z, pixelWidth, pixelHeight, colorValue, Rotation, Technique
        Else
            Draw_Texture CurrentGrhIndex, X, Y, Z, .pixelWidth, .pixelHeight, colorValue, Rotation, Technique

        End If

    End With
    
    Exit Sub

error:
    'MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & vbCrLf & err.Description, vbExclamation, "[ " & err.Number & " ] Error"

End Sub

'
' Dibujamos un Grh a escala de 64x64
'
Public Sub Draw_Grh_Menu(ByRef grh As grh, _
                         ByVal X As Integer, _
                         ByVal Y As Integer, _
                         ByVal Z As Single, _
                         ByVal Center As Byte, _
                         Optional ByVal Animate As Byte = 0, _
                         Optional ByVal killAtEnd As Byte = 0, _
                         Optional ByVal Colour As Long = -1, _
                         Optional ByVal Rotation As Single = 0, _
                         Optional ByVal Technique As Integer = eTechnique.t_Default, _
                         Optional ByVal pixelWidth As Integer = 0, _
                         Optional ByVal pixelHeight As Integer = 0, _
                         Optional ByVal FxSlot As Byte = 0, _
                         Optional NotBig As Boolean = False, _
                         Optional ByVal SoloBig As Boolean = False)
    
    Dim CurrentGrhIndex As Long

    On Error GoTo error

      If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If Animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.Speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    Dim W    As Long

    Dim H    As Long

    Dim Mult As Long

    #If ModoBig = 0 Then
        W = 32
        H = 32
        Mult = 1
        Mult = 1
    #Else
        W = 64
        H = 64
        Mult = 2
        Mult = 2
    #End If
    
    With GrhData(CurrentGrhIndex)
        
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * W * 0.5) + W * 0.5

            End If
                
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * H) + H

            End If

        End If

        If pixelWidth <> 0 And pixelHeight <> 0 Then
            Draw_Texture CurrentGrhIndex, X, Y, Z, pixelWidth, pixelHeight, Colour, Rotation, Technique
        Else
            Draw_Texture CurrentGrhIndex, X, Y, Z, .pixelWidth, .pixelHeight, Colour, Rotation, Technique

        End If

    End With
    
    Exit Sub

error:

        'MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & vbCrLf & err.Description, vbExclamation, "[ " & err.Number & " ] Error"

End Sub
'
' Dibujamos una textura
'
Public Sub Draw_Texture(ByVal GrhIndex As Long, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal Depth As Single, _
                        ByVal Width As Long, _
                        ByVal Height As Long, _
                        ByVal Colour As Long, _
                        ByVal Rotation As Single, _
                        ByVal Technique As eTechnique, _
                        Optional ByVal NotBig As Boolean = False)

        '<EhHeader>
        On Error GoTo Draw_Texture_Err

        '</EhHeader>
                   
        Dim destination As wGL_Rectangle, Source As wGL_Rectangle

        Dim data()      As Byte
    
        On Error GoTo ErrHandler
    
100     With GrhData(GrhIndex)

            If .Alpha Then
                Technique = eTechnique.t_Alpha

            End If
             
            If .FileNum = 5443 Then
                Debug.Print .FileNum
            End If
            
106         If g_Textures(.FileNum) = 0 Then
              
              
                #If Testeo = 1 Then
                    
                    #If ModoBig = 0 Then
                        data = LoadBytes("resource\pngs1\" & CStr(.FileNum) & ".PNG")
                    #Else
                        data = LoadBytes("resource\pngs2\" & CStr(.FileNum) & ".PNG")
                    #End If

                #Else
                    Call Get_Image(DirGraficos & GRH_RESOURCE_FILE, CStr(.FileNum), data)
                #End If
                    
                '#End If
                
                Dim Material As Integer

110             Material = wGL_Graphic_Renderer.Create_Material
112             g_Textures(.FileNum) = Material
114             Call wGL_Graphic_Renderer.Update_Material_Texture(Material, 0, wGL_Graphic.Create_Texture_From_Image(data))
                
            End If
        
116         destination.x1 = X: destination.y1 = Y: destination.X2 = X + Width: destination.Y2 = Y + Height
        
118         Call wGL_Graphic_Renderer.Draw(destination, .src, Depth, Rotation, Colour, g_Textures(.FileNum), Techniques(Technique))

        End With
    
        Exit Sub

ErrHandler:
    
        '<EhFooter>
        Exit Sub

Draw_Texture_Err:
        LogError err.Description & vbCrLf & "in Draw_Texture " & "at line " & Erl

        '</EhFooter>
End Sub

Public Sub Draw_Texture_Default(ByVal GrhIndex As Long, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal Depth As Single, _
                        ByVal Width As Long, _
                        ByVal Height As Long, _
                        ByVal Colour As Long, _
                        ByVal Rotation As Single, _
                        ByVal Technique As eTechnique)

        '<EhHeader>
        On Error GoTo Draw_Texture_Err

        '</EhHeader>
                   
        Dim destination As wGL_Rectangle, Source As wGL_Rectangle

        Dim data()      As Byte
    
        On Error GoTo ErrHandler
    
100     With GrhDataDefault(GrhIndex)

            If .Alpha Then
                Technique = eTechnique.t_Alpha

            End If
             
            
106         If g_Textures(.FileNum) = 0 Then
              
                Call Get_Image(DirGraficos & GRH_RESOURCE_FILE_DEFAULT, CStr(.FileNum), data)
                
                Dim Material As Integer

110             Material = wGL_Graphic_Renderer.Create_Material
112             g_Textures(.FileNum) = Material
114             Call wGL_Graphic_Renderer.Update_Material_Texture(Material, 0, wGL_Graphic.Create_Texture_From_Image(data))
                
            End If
        
116         destination.x1 = X: destination.y1 = Y: destination.X2 = X + Width: destination.Y2 = Y + Height
        
118         Call wGL_Graphic_Renderer.Draw(destination, .src, Depth, Rotation, Colour, g_Textures(.FileNum), Techniques(Technique))

        End With
    
        Exit Sub

ErrHandler:
    
        '<EhFooter>
        Exit Sub

Draw_Texture_Err:
        LogError err.Description & vbCrLf & "in Draw_Texture " & "at line " & Erl

        '</EhFooter>
End Sub

Public Sub Draw_Texture_Graphic(ByVal FileNum As Long, _
                                ByVal X As Integer, _
                                ByVal Y As Integer, _
                                ByVal Depth As Single, _
                                ByVal Width As Integer, _
                                ByVal Height As Integer, _
                                ByVal sX As Single, _
                                ByVal sY As Single, _
                                ByVal pixelWidth As Integer, _
                                ByVal pixelHeight As Integer, _
                                ByVal Colour As Long, _
                                ByVal Rotation As Single, _
                                ByVal Technique As Integer)

        '<EhHeader>
        On Error GoTo Draw_Texture_Graphic_Err

        '</EhHeader>
                   
        Dim destination As wGL_Rectangle, Source As wGL_Rectangle
                   
        Dim sX1         As Single, sX2 As Single

        Dim sY1         As Single, sY2 As Single

        Dim data()      As Byte

        Dim Material    As Integer
    
        On Error GoTo ErrHandler

100     If g_Textures(FileNum) = 0 Then

            
            Call Get_Image(DirGraficos & GRH_RESOURCE_FILE, CStr(FileNum), data)

        
104         Material = wGL_Graphic_Renderer.Create_Material
106         g_Textures(FileNum) = Material
   
108         Call wGL_Graphic_Renderer.Update_Material_Texture(Material, 0, wGL_Graphic.Create_Texture_From_Image(data))

        End If
        
110     sX1 = sX / Width
112     sY1 = sY / Height
    
114     destination.x1 = X: destination.y1 = Y: destination.X2 = X + Width: destination.Y2 = Y + Height
116     Source.x1 = sX1: Source.y1 = sY1: Source.X2 = sY1 + pixelWidth / Width: Source.Y2 = sX1 + pixelHeight / Height
        
118     Call wGL_Graphic_Renderer.Draw(destination, Source, Depth, Rotation, Colour, g_Textures(FileNum), Techniques(Technique))
    
        Exit Sub
ErrHandler:
    
        '<EhFooter>
        Exit Sub

Draw_Texture_Graphic_Err:
        LogError err.Description & vbCrLf & "in Draw_Texture_Graphic " & "at line " & Erl

        '</EhFooter>
End Sub

Public Sub Draw_Texture_Graphic_Gui(ByVal FileNum As Long, _
                                    ByVal X As Integer, _
                                    ByVal Y As Integer, _
                                    ByVal Depth As Single, _
                                    ByVal Width As Integer, _
                                    ByVal Height As Integer, _
                                    ByVal sX As Single, _
                                    ByVal sY As Single, _
                                    ByVal pixelWidth As Integer, _
                                    ByVal pixelHeight As Integer, _
                                    ByVal Colour As Long, _
                                    ByVal Rotation As Single, _
                                    ByVal Technique As eTechnique)

        '<EhHeader>
        On Error GoTo Draw_Texture_Graphic_Err

        '</EhHeader>
                   
        Dim destination As wGL_Rectangle, Source As wGL_Rectangle
                   
        Dim sX1         As Single, sX2 As Single

        Dim sY1         As Single, sY2 As Single

        Dim data()      As Byte

        Dim Material    As Integer
        
        On Error GoTo ErrHandler

100     If g_Textures_Gui(FileNum) = 0 Then
            data = LoadBytes("\resource\interface\gui\" & CStr(FileNum) & ".png")
104         Material = wGL_Graphic_Renderer.Create_Material
106         g_Textures_Gui(FileNum) = Material
108         Call wGL_Graphic_Renderer.Update_Material_Texture(Material, 0, wGL_Graphic.Create_Texture_From_Image(data))

        End If
        
110     sX1 = sX / Width
112     sY1 = sY / Height
    
114      destination.x1 = X: destination.y1 = Y: destination.X2 = X + Width: destination.Y2 = Y + Height
116     Source.x1 = sX1: Source.y1 = sY1: Source.X2 = sY1 + pixelWidth / Width: Source.Y2 = sX1 + pixelHeight / Height
        
118     Call wGL_Graphic_Renderer.Draw(destination, Source, Depth, Rotation, Colour, g_Textures_Gui(FileNum), Techniques(Technique))
    
        Exit Sub
ErrHandler:
    
        '<EhFooter>
        Exit Sub

Draw_Texture_Graphic_Err:
        LogError err.Description & vbCrLf & "in Draw_Texture_Graphic " & "at line " & Erl

        '</EhFooter>
End Sub

Public Sub Draw_Avatar(ByVal FileNum As Long, _
                                ByVal X As Integer, _
                                ByVal Y As Integer, _
                                ByVal Depth As Single, _
                                ByVal Width As Integer, _
                                ByVal Height As Integer, _
                                ByVal sX As Single, _
                                ByVal sY As Single, _
                                ByVal pixelWidth As Integer, _
                                ByVal pixelHeight As Integer, _
                                ByVal Colour As Long, _
                                ByVal Rotation As Single, _
                                ByVal Technique As eTechnique)
        '<EhHeader>
        On Error GoTo Draw_Avatar_Err
        '</EhHeader>
                   
        Dim destination As wGL_Rectangle, Source As wGL_Rectangle
                   
        Dim sX1         As Single, sX2 As Single
        Dim sY1         As Single, sY2 As Single
        Dim data()      As Byte
        Dim Material    As Integer
        
        On Error GoTo ErrHandler
100     If g_Textures_Avatars(FileNum) = 0 Then
              Call Get_Image(DirGraficos & AVATARS_RESOURCE_FILE, CStr(FileNum), data, True)
104         Material = wGL_Graphic_Renderer.Create_Material
106         g_Textures_Avatars(FileNum) = Material
108         Call wGL_Graphic_Renderer.Update_Material_Texture(Material, 0, wGL_Graphic.Create_Texture_From_Image(data))
        End If
        
110     sX1 = sX / Width
112     sY1 = sY / Height
    
114     destination.x1 = X: destination.y1 = Y: destination.X2 = X + Width: destination.Y2 = Y + Height
116     Source.x1 = sX1: Source.y1 = sY1: Source.X2 = sY1 + pixelWidth / Width: Source.Y2 = sX1 + pixelHeight / Height
        
118     Call wGL_Graphic_Renderer.Draw(destination, Source, Depth, Rotation, Colour, g_Textures_Avatars(FileNum), Techniques(Technique))
    
        Exit Sub
ErrHandler:
    
        '<EhFooter>
        Exit Sub

Draw_Avatar_Err:
        LogError err.Description & vbCrLf & _
           "in Draw_Avatar " & _
           "at line " & Erl

        '</EhFooter>
End Sub

Public Sub Draw_Texture_Graphic_MiniMap(ByVal FileNum As Long, _
   ByVal X As Integer, _
   ByVal Y As Integer, _
   ByVal Depth As Single, _
   ByVal Width As Integer, _
   ByVal Height As Integer, _
   ByVal sX As Single, _
   ByVal sY As Single, _
   ByVal pixelWidth As Integer, _
   ByVal pixelHeight As Integer, _
   ByVal Colour As Long, _
   ByVal Rotation As Single, _
   ByVal Technique As eTechnique)

        '<EhHeader>
        On Error GoTo Draw_Texture_Graphic_Err

        '</EhHeader>
                   
        Dim destination As wGL_Rectangle, Source As wGL_Rectangle
                   
        Dim sX1         As Single, sX2 As Single

        Dim sY1         As Single, sY2 As Single

        Dim data()      As Byte

        Dim Material    As Integer

        Dim filePath As String
        
        On Error GoTo ErrHandler
        
100     If g_Textures_MiniMapa(FileNum) = 0 Then
            data = LoadBytes(Replace(MiniMap_FilePath, App.path, vbNullString) & CStr(FileNum) & ".png")
104         Material = wGL_Graphic_Renderer.Create_Material
106         g_Textures_MiniMapa(FileNum) = Material
108         Call wGL_Graphic_Renderer.Update_Material_Texture(Material, 0, wGL_Graphic.Create_Texture_From_Image(data))

        End If
        
110     sX1 = sX / Width
112     sY1 = sY / Height
    
114     destination.x1 = X: destination.y1 = Y: destination.X2 = X + Width: destination.Y2 = Y + Height
116     Source.x1 = sX1: Source.y1 = sY1: Source.X2 = sY1 + pixelWidth / Width: Source.Y2 = sX1 + pixelHeight / Height
        
118     Call wGL_Graphic_Renderer.Draw(destination, Source, Depth, Rotation, Colour, g_Textures_MiniMapa(FileNum), Techniques(Technique))
    
        Exit Sub
ErrHandler:
    
        '<EhFooter>
        Exit Sub

Draw_Texture_Graphic_Err:
        LogError err.Description & vbCrLf & _
           "in Draw_Texture_Graphic " & _
           "at line " & Erl

        '</EhFooter>
End Sub

' Dibujamos un Texto
'

Public Sub Draw_Text(ByVal Font As eFonts, _
                     ByVal Size As Integer, _
                     ByVal X As Single, _
                     ByVal Y As Single, _
                     ByVal Depth As Single, _
                     ByVal Blur As Single, _
                     ByVal Colour As Long, _
                     ByVal Alineation As wGL_Graphic_Font_Alignment, _
                     ByVal Text As String, _
                     ByVal Shadow As Boolean, _
                     Optional ByVal NotBig As Boolean = False)

    If Text = vbNullString Then Exit Sub
    
    
    #If ModoBig > 0 Then
        If NotBig = False Then
            Size = (Size * 2)
        End If
      '  Font = eFonts.f_Booter
    #End If
    
    
    If Shadow Then
        Call wGL_Graphic_Renderer.Draw_Text(Font, Size, X - 1, Y, Depth - 0.0000001, &HFF212121, Alineation, Text)
        Call wGL_Graphic_Renderer.Draw_Text(Font, Size, X + 1, Y, Depth - 0.0000001, &HFF212121, Alineation, Text)
        Call wGL_Graphic_Renderer.Draw_Text(Font, Size, X, Y - 1, Depth - 0.0000001, &HFF212121, Alineation, Text)
        Call wGL_Graphic_Renderer.Draw_Text(Font, Size, X, Y + 1, Depth - 0.0000001, &HFF212121, Alineation, Text)
    End If

    Call wGL_Graphic_Renderer.Draw_Text(Font, Size, X, Y, Depth, Colour, Alineation, Text)
End Sub

Function ShowNextFrame() As Boolean

    '***************************************************
    'Author: Arron Perkins
    'Last Modification: 08/14/07
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Updates the game's model and renders everything.
    '***************************************************

    If Not EngineRun Then Exit Function

    If ClientSetup.bFps = 0 Then ClientSetup.bFps = 255
    
      If ((GetTickCount() - FrameTime >= TARGET_FPS_MS / ClientSetup.bFps) Or (ClientSetup.bFps = 1)) Or (ClientSetup.bFps = 2) Then
        
        If Not MirandoCuenta Then
            Call wGL_Graphic.Use_Device(&H0)
            Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, &H0, 1#, 0)
            
            Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmMain.MainViewPic.ScaleWidth, FrmMain.MainViewPic.ScaleHeight)

            
            If (ModoNoche) Then
                Call SetAmbientColor(96, 96, 96, 255)
            End If
            
            If Not UserCiego Then
                Call RenderScreen
            End If
            
            Call Dialogos.Render
            Call DialogosClanes.Draw
    
            If (ModoNoche) Then
                Call SetAmbientColor(255, 255, 255, 255)
            End If
            
            Call wGL_Graphic_Renderer.Flush
        End If
        
        FrameTime = GetTickCount
        timerElapsedTime = GetElapsedTime
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

        
        ShowNextFrame = True
    End If

End Function

Private Function CharIsGuild(ByVal Char As Integer) As Boolean

    Dim tempTag   As String

    Dim tempPos   As Integer

    Dim miTag     As String

    Dim miTempPos As Integer
    
    If Char = UserCharIndex Then Exit Function
    
    With CharList(Char)
    
        miTempPos = getTagPosition(CharList(UserCharIndex).Nombre)
        miTag = mid$(CharList(UserCharIndex).Nombre, miTempPos)
        tempPos = getTagPosition(.Nombre)
        tempTag = mid$(.Nombre, tempPos)
        
        If tempTag = miTag And miTag <> vbNullString And tempTag <> vbNullString And GuildLevel >= 12 Then
            CharIsGuild = True

            Exit Function

        End If
        
        CharIsGuild = False
    End With

End Function

Private Sub CharRender(ByVal CharIndex As Long, _
                       ByVal PixelOffsetX As Integer, _
                       ByVal PixelOffsetY As Integer, _
                       ByVal Mult As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 25/05/2011 (Amraphen)
    'Draw char's to screen without offcentering them
    '16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
    '25/05/2011: Amraphen - Agregado movimiento de armas al golpear.
    '***************************************************1
    Dim moved    As Boolean

    Dim attacked As Boolean

    Dim Pos      As Integer

    Dim line     As String

    Dim Color    As Long
        
    Dim A        As Long

    With CharList(CharIndex)
        
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding

                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding

                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If
            
            If .scrollDirectionX = 0 And .scrollDirectionY = 0 Then
                .Moving = False

            End If
            
            Call Audio.UpdateSource(.SoundSource, .Pos.X, .Pos.Y)

        Else

            If .Muerto Then
                ' If CharIndex <> UserCharIndex Then

                ' Si no somos nosotros, esperamos un intervalo
                ' antes de poner la animación idle para evitar saltos
                'If FrameTime - .LastStep > TIME_CASPER_IDLE Then
                ' .Body = BodyData(CASPER_BODY_IDLE)
                ' .Body.Walk(.Heading).started = FrameTime
                '  .Idle = True

                ' End If
                    
                'Else
                ' .Body = BodyData(CASPER_BODY_IDLE)
                '  .Body.Walk(.Heading).started = FrameTime
                '  .Idle = True

                '   End If

            Else

                'Stop animations
                If .Navegando = False Then
                    .Body.Walk(.Heading).started = 0

                    If Not .MovArmaEscudo Then
                        .Arma.WeaponWalk(.Heading).started = 0
                        .Escudo.ShieldWalk(.Heading).started = 0

                    End If

                End If

            End If

        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim PosX As Integer, PosY As Integer

        PosX = .Pos.X
        PosY = .Pos.Y
        
        ' Npcs sin cabeza
            
        If .NpcIndex > 0 Then
            If Abs(FrmMain.tX - .Pos.X) <= 2 And (Abs(FrmMain.tY - .Pos.Y)) <= 2 Then
                If Abs(FrmMain.tX - .Pos.X) = 0 And (Abs(FrmMain.tY - .Pos.Y)) = 0 Then
                    If NpcIndex_MouseHover <> .NpcIndex Then
                        NpcIndex_MouseHover = .NpcIndex
                        CharIndex_MouseHover = CharIndex
                    End If
                Else
                    NpcIndex_MouseHover = 0
                End If
                
                If .ColorNick = eNickColor.ieCastleGuild Then
                    Color = ARGB(249, 216, 150, 255)
                ElseIf .ColorNick = eNickColor.ieCastleUser Then
                    Color = ARGB(245, 144, 54, 255)
                Else
                    Color = ARGB(230, 185, 7, 255)
    
                End If
                        
                Pos = getTagPosition(.Nombre)
    
                line = Left$(.Nombre, Pos - 2)
                          
                Call Draw_Text(eFonts.f_Tahoma, 14, PixelOffsetX + (16 * Mult), PixelOffsetY + (30 * Mult), To_Depth(3, PosX, PosY, 8), 0, Color, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, line, True)
                
                line = mid$(.Nombre, Pos)
                Color = ARGB(240, 200, 0, 255)
                Call Draw_Text(eFonts.f_Tahoma, 14, PixelOffsetX + (16 * Mult), PixelOffsetY + (45 * Mult), To_Depth(3, PosX, PosY, 8), 0, Color, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, line, True)
            Else

                If .NpcIndex = NpcIndex_MouseHover Then
                    NpcIndex_MouseHover = 0

                End If

            End If
                
        End If
        
        Dim Technique   As eTechnique

        Dim CharIsValid As Boolean
        
        If (.Invisible) Then

            Dim ValidFaction As Boolean

            Dim ValidGroup   As Boolean
            
            If UserCharIndex <> CharIndex Then
                If .ColorNick = eNickColor.ieCAOS Or .ColorNick = eNickColor.ieArmada Then
                    If CharList(UserCharIndex).ColorNick = .ColorNick Then
                        ValidFaction = False

                    End If

                End If
                
                If .GroupIndex > 0 Then
                    If CharList(UserCharIndex).GroupIndex = .GroupIndex Then
                        ValidGroup = True
                    Else
                        ValidGroup = False

                    End If

                Else
                    ValidGroup = False

                End If

            End If
                
            If (CharIsGuild(CharIndex) Or ValidFaction) Or (.Intermitencia) Or (ValidGroup) Or ((UserCharIndex = CharIndex) And ClientSetup.bConfig(eSetupMods.SETUP_PERSONAJEOCULTOENINVI) = 0) Then
                Color = ARGB(255, 255, 255, 100)
                CharIsValid = True
                Technique = t_Alpha
            Else
                CharIsValid = False
                Technique = t_Default

            End If
            
        Else
            Color = ARGB(255, 255, 255, 255)
            Technique = t_Default
            'If (CharIndex = UserCharIndex) Or (Not .Invisible) Then
            CharIsValid = True

            'Else
            'CharIsValid = False
            'End If
        End If
            
        Dim View As Boolean
                        
        View = True

        'If CharList(UserCharIndex).Muerto And Not UserCharIndex = CharIndex Then
        'If .Muerto = 0 And Not .IsNpc Then
        'View = False

        'End If

        'End If
            
        If View Then

            If UserEnvenenado And UserCharIndex = CharIndex Then
                Color = ARGB(90, 235, 50, 200)

            End If
            
            If CharIsValid Then

                ' Aura Global
                'If .Aura(5).Walk(.Heading).GrhIndex Then Call Draw_Grh(.Aura(4).Walk(.Heading), PixelOffsetX, PixelOffsetY - 15, To_Depth(3, PosX, PosY, 1), 1, 1, 0, Color, , eTechnique.t_Alpha)
                
                'If .Heading = NORTH Then
                'call Draw_Grh(.Aura.Walk(.Heading), PixelOffsetX, PixelOffsetY - 15, To_Depth(3, PosX, PosY, 3), 1, 1, 0, Color, , eTechnique.t_Alpha)
                ' Else
                'Call Draw_Grh(.Aura.Walk(.Heading), PixelOffsetX, PixelOffsetY - 15, To_Depth(3, PosX, PosY, 1), 1, 1, 0, Color, , eTechnique.t_Alpha)

                'End If

                ' End If
                    
                ' Modo Streamer
                ' If .Streamer Then
                'Call Draw_Texture_Graphic_Gui(88, PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 9), 20, 20, 0, 0, 20, 20, Color, 0, Technique)
                'End If
            
                If .BodyAttack.Walk(.Heading).GrhIndex Then
                    'If AttackNpc Then
                    'Call Draw_Grh(.BodyAttack.Walk(.Heading), PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 2), 1, 1, 0, Color, , Technique)
                    'Else

                    If .Body.Walk(.Heading).GrhIndex Then
                        Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 2), 1, 1, 0, Color, , Technique)

                    End If

                    'End If

                Else

                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex Then
                        If .NpcIndex > 0 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + .Body.BodyOffSet(.Heading).X, PixelOffsetY + .Body.BodyOffSet(.Heading).Y, To_Depth(3, PosX, PosY, 3), 1, 1, 0, Color, , eTechnique.t_Alpha)
                        Else

                            If .Muerto Then
                                Color = ARGB(255, 255, 255, 150)
                                
                            Else
                                Color = ARGB(255, 255, 255, 255)

                            End If
                            
                            ' Aura de la Armadura
                            If .Aura(1).Walk(.Heading).GrhIndex Then
                                Call Draw_Grh(.Aura(1).Walk(.Heading), PixelOffsetX + .Body.BodyOffSet(.Heading).X, PixelOffsetY + .Body.BodyOffSet(.Heading).Y + 25, To_Depth(3, PosX, PosY, 1), 1, 1, 0, .Aura(1).Color, , eTechnique.t_Alpha)

                            End If

                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + .Body.BodyOffSet(.Heading).X, PixelOffsetY + .Body.BodyOffSet(.Heading).Y, To_Depth(3, PosX, PosY, 3), 1, 1, 0, Color, , eTechnique.t_Alpha)

                        End If

                    End If
                    
                End If
                    
                Dim MultValue As Integer
                    
                If Mult = 2 Then
                    MultValue = -1

                End If
                    
                'Draw Head
                '.Head.Head(.Heading).GrhIndex = 0
                If .Head.Head(.Heading).GrhIndex Then
                    Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, To_Depth(3, PosX, PosY, 4), 1, 0, , Color, , eTechnique.t_Alpha)

                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then

                        ' Aura del Casco
                        If .Aura(3).Walk(.Heading).GrhIndex Then Call Draw_Grh(.Aura(3).Walk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 25, To_Depth(3, PosX, PosY, 1), 1, 1, 0, .Aura(3).Color, , eTechnique.t_Alpha)

                        Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + ((.Body.HeadOffset.Y - .OffsetY)), To_Depth(3, PosX, PosY, 5), 1, 0, , Color, , eTechnique.t_Alpha)

                    End If

                    ' Arma
                    
                    If .Heading = NORTH Or .Heading = WEST Then
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then

                            ' Aura del Arma
                            If .Aura(2).Walk(.Heading).GrhIndex Then Call Draw_Grh(.Aura(2).Walk(.Heading), PixelOffsetX, PixelOffsetY + 25, To_Depth(3, PosX, PosY, 1), 1, 1, 0, .Aura(2).Color, , eTechnique.t_Alpha)
                
                            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 1), 1, 1, 0, Color, , eTechnique.t_Alpha)

                        End If
                        
                    Else

                        If .Arma.WeaponWalk(.Heading).GrhIndex Then

                            ' Aura del Arma
                            If .Aura(2).Walk(.Heading).GrhIndex Then Call Draw_Grh(.Aura(2).Walk(.Heading), PixelOffsetX, PixelOffsetY + 25, To_Depth(3, PosX, PosY, 1), 1, 1, 0, .Aura(2).Color, , eTechnique.t_Alpha)
                            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 5), 1, 1, 0, Color, , eTechnique.t_Alpha)

                        End If
                        
                    End If
                
                    ' Escudo
                    If .Heading = NORTH Or .Heading = EAST Then

                        ' Aura del Escudo
                        If .Aura(4).Walk(.Heading).GrhIndex Then Call Draw_Grh(.Aura(4).Walk(.Heading), PixelOffsetX, PixelOffsetY + 25, To_Depth(3, PosX, PosY, 1), 1, 1, 0, .Aura(4).Color, , eTechnique.t_Alpha)
                
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 1), 1, 1, 0, Color, , eTechnique.t_Alpha)
                    Else

                        ' Aura del Escudo
                        If .Aura(4).Walk(.Heading).GrhIndex Then Call Draw_Grh(.Aura(4).Walk(.Heading), PixelOffsetX, PixelOffsetY + 25, To_Depth(3, PosX, PosY, 1), 1, 1, 0, .Aura(4).Color, , eTechnique.t_Alpha)
                
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, To_Depth(3, PosX, PosY, 6), 1, 1, 0, Color, , eTechnique.t_Alpha)

                    End If
                        
                    If ControlActivated Then

                        ' Hp
                        If .MaxHp <> 0 Then
                        
                            Draw_Texture 24773, PixelOffsetX - 10, PixelOffsetY + 45, To_Depth(3, PosX, PosY, 8), GrhData(24773).pixelWidth, GrhData(24773).pixelHeight, Color, 0, eTechnique.t_Default
                            
                            If .MinHp > 0 Then
                                Draw_Texture 24772, PixelOffsetX - 10, PixelOffsetY + 45, To_Depth(3, PosX, PosY, 9), (((.MinHp / 100) / (.MaxHp / 100)) * GrhData(24772).pixelWidth), GrhData(24772).pixelHeight, -1, 0, 0

                            End If

                        End If
                            
                        ' Man
                        If .MinMan <> 0 Then
                            Draw_Texture 24774, PixelOffsetX - 10, PixelOffsetY + 60, To_Depth(3, PosX, PosY, 8), GrhData(24774).pixelWidth, GrhData(24774).pixelHeight, -1, 0, 0
                            Draw_Texture 24771, PixelOffsetX - 10, PixelOffsetY + 60, To_Depth(3, PosX, PosY, 9), (((.MinMan / 100) / (.MaxMan / 100)) * GrhData(24771).pixelWidth), GrhData(24771).pixelHeight, -1, 0, 0

                        End If

                    End If
                        
                    ' Barritas de BAR sobre la cabeza del usuario
                    Dim Y_BAR As Integer

                    Dim Value As Long
                        
                    Y_BAR = 20
                        
                    If .BarMax <> 0 Then
                        Value = (((.BarMax / 10) * GrhData(BAR_BACKGROUND).pixelWidth) - ((.BarMin / 10)) * GrhData(BAR_BACKGROUND).pixelWidth)
                        Value = GrhData(BAR_BACKGROUND).pixelWidth - (((.BarMin / 100) / (.BarMax / 100)) * GrhData(BAR_BACKGROUND).pixelWidth)
                        
                        Draw_Texture BAR_BORDER, PixelOffsetX, PixelOffsetY - (35 * Mult), To_Depth(3, PosX, PosY, 8), GrhData(BAR_BORDER).pixelWidth, GrhData(BAR_BORDER).pixelHeight, -1, 0, 0
                        Draw_Texture BAR_BACKGROUND, PixelOffsetX + (2 * Mult), PixelOffsetY - (33 * Mult), To_Depth(3, PosX, PosY, 9), Value, GrhData(BAR_BACKGROUND).pixelHeight, -1, 0, 0
                          
                    End If
                
                End If
            
                'Draw name over head
                If LenB(.Nombre) > 0 And .Body.Walk(.Heading).GrhIndex And Not .NpcIndex > 0 Then
                    If Nombres Then
                        Pos = getTagPosition(.Nombre)
                                    
                        If .Priv = 0 Then
                            If .Criminal Then
                                Color = ARGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b, 255)
                            Else
                                Color = ARGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b, 255)

                            End If
                            
                            ' Colores nuevos
                            If .ColorNick = eNickColor.ieCastleGuild Then
                                Color = ARGB(249, 216, 150, 255)
                            ElseIf .ColorNick = eNickColor.ieCastleUser Then
                                Color = ARGB(245, 144, 54, 255)
                            ElseIf .ColorNick = eNickColor.ieShield Then
                                Color = ARGB(245, 40, 245, 255)
                            ElseIf .ColorNick = eNickColor.ieAtacable Then
                                Color = ARGB(160, 240, 187, 255)

                            End If
                            
                        Else
                            Color = ARGB(ColoresPJ(.Priv).r, ColoresPJ(.Priv).g, ColoresPJ(.Priv).b, 255)

                        End If
                        
                        'Color = Color Or ((255 And Not &H80) * &H1000000) Or &H80000000
                        
                        'Nick
                        line = Left$(.Nombre, Pos - 2)
                        
                        If .Invisible Then
                            Color = ARGB(240, 240, 240, 200)
                        Else

                            If .Priv = 0 Then

                                Select Case .ColorNick

                                    Case eNickColor.ieArmada
                                        Color = ARGB(8, 88, 167, 255)

                                    Case eNickColor.ieCAOS
                                        Color = ARGB(176, 8, 8, 255)

                                    Case eNickColor.ieCriminal
                                        
                                End Select

                            End If

                        End If
                        
                        Call Draw_Text(f_Chat, 14, PixelOffsetX + (16 * Mult), PixelOffsetY + (30 * Mult), To_Depth(3, PosX, PosY, 8), 0, Color, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, line, True)
                    
                        line = mid$(.Nombre, Pos)
                        
                        Call Draw_Text(f_Chat, 14, PixelOffsetX + (16 * Mult), PixelOffsetY + (45 * Mult), To_Depth(3, PosX, PosY, 8), 0, Color, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, line, True)
          
                    End If
                    
                End If
                
            End If

        End If
        
        'Update dialogs
        'Call Dialogos.UpdateDialogPos(PixelOffsetX + 20, PixelOffsetY + .Body.HeadOffset.Y + 10, CharIndex)
        
        'not muerto + npc = no ve dialogos
        If Not (CharList(CharIndex).Muerto = False And CharList(UserCharIndex).Muerto And Not .NpcIndex > 0) Then
            'Update dialogs
            Call Dialogos.UpdateDialogPos(PixelOffsetX + TilePixelWidth \ 2, PixelOffsetY + (.Body.HeadOffset.Y) + 10, 0#, CharIndex)  '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo

        End If
                
           Dim Colour As Long

            Colour = ARGB(255, 255, 255, ClientSetup.bAlpha)
            
        ' Draw Meditation
        If .FxIndex <> 0 And .fX.started <> 0 Then
            Call Draw_GrhFX(CharIndex, 0, .fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + (FxData(.FxIndex).OffsetY * Mult), To_Depth(3, PosX, PosY, 10), 1, 1, 1, Colour, , eTechnique.t_Alpha)
        End If
        
        
        Dim I As Long
        
        ' Draw Effects
        If .FxCount > 0 Then

            For I = 1 To .FxCount

                If .FxList(I).FxIndex > 0 And .FxList(I).started <> 0 Then
    
                    Call Draw_GrhFX(CharIndex, I, .FxList(I), PixelOffsetX + FxData(.FxList(I).FxIndex).OffsetX, PixelOffsetY + (FxData(.FxList(I).FxIndex).OffsetY * Mult), To_Depth(3, PosX, PosY, 10 + A), 1, 1, 1, Colour, , eTechnique.t_Alpha)

                End If

                If .FxList(I).started = 0 Then
                    .FxList(I).FxIndex = 0

                End If

            Next I

            If .FxList(.FxCount).started = 0 Then
                .FxCount = .FxCount - 1

            End If

        End If
        


    End With

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
    
    On Error GoTo SetCharacterFx_Err
    

    If fX = 0 Then Exit Sub

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    Dim indice As Byte

    With CharList(CharIndex)
    
        indice = Char_FX_Group_Next_Open(CharIndex)
        
        .FxList(indice).FxIndex = fX
        
        Call InitGrh(.FxList(indice), FxData(fX).Animacion, , Loops)
            
    End With

    
    Exit Sub

SetCharacterFx_Err:
    Resume Next
    
End Sub
Public Function Char_FX_Group_Next_Open(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LoopC As Long
    
    If CharList(char_index).FxCount = 0 Then
        CharList(char_index).FxCount = 1
        ReDim CharList(char_index).FxList(1 To 1)
        Char_FX_Group_Next_Open = 1
        Exit Function

    End If
    
    LoopC = 1

    Do Until CharList(char_index).FxList(LoopC).FxIndex = 0

        If LoopC = CharList(char_index).FxCount Then
            Char_FX_Group_Next_Open = CharList(char_index).FxCount + 1
            CharList(char_index).FxCount = Char_FX_Group_Next_Open
            ReDim Preserve CharList(char_index).FxList(1 To Char_FX_Group_Next_Open)
            Exit Function

        End If

        LoopC = LoopC + 1
    Loop

    Char_FX_Group_Next_Open = LoopC
    Exit Function

ErrorHandler:
    CharList(char_index).FxCount = 1
    ReDim CharList(char_index).FxList(1 To 1)
    Char_FX_Group_Next_Open = 1

End Function
Public Sub SetCharacterFxMap(ByVal X As Integer, _
                             ByVal Y As Integer, _
                             ByVal fX As Integer, _
                             ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    With MapData(X, Y)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If

    End With

End Sub

' RTREE
Public Function GetCharacterDimension(ByVal CharIndex As Integer, _
                                      ByRef RangeX As Single, _
                                      ByRef RangeY As Single)
        '<EhHeader>
        On Error GoTo GetCharacterDimension_Err
        '</EhHeader>

100     With CharList(CharIndex)
    
102         If (.iBody <> 0) Then
104             RangeX = GrhData(.Body.Walk(.Heading).GrhIndex).TileWidth
106             RangeY = GrhData(.Body.Walk(.Heading).GrhIndex).TileHeight
            End If
            
108         If (.iHead <> 0) Then
110             If (GrhData(.Head.Head(.Heading).GrhIndex).TileWidth > RangeX) Then
112                 RangeX = GrhData(.Head.Head(.Heading).GrhIndex).TileWidth
                End If
                
114             RangeY = RangeY + GrhData(.Head.Head(.Heading).GrhIndex).TileHeight + 2# 'Name + Guild
            End If
            
        End With
        
        '<EhFooter>
        Exit Function

GetCharacterDimension_Err:
        LogError err.Description & vbCrLf & _
           "in GetCharacterDimension " & _
           "at line " & Erl

        '</EhFooter>
End Function

'######################
'######################
'                       Procedimientos que no reciben modificaciones
'######################
'######################

Public Function ARGB(ByVal red As Long, ByVal green As Long, ByVal blue As Long, ByVal Alpha As Long) As Long
    If Alpha > 127 Then
       ARGB = ((Alpha - 128) * &H1000000 Or &H80000000) Or blue Or (green * &H100&) Or (red * &H10000)
    Else
       ARGB = (Alpha * &H1000000) Or blue Or (green * &H100&) Or (red * &H10000)
    End If
End Function

Public Function To_Depth(ByVal Layer As Single, _
                         Optional ByVal X As Single = 1, _
                         Optional ByVal Y As Single = 1, _
                         Optional ByVal Z As Single = 1) As Single

    To_Depth = -1# + (Layer * 0.1) + ((Y - 1) * 0.001) + ((X - 1) * 0.00001) + ((Z - 1) * 0.000001)
    
End Function

Public Function LoadBytes(ByVal FileName As String) As Byte()
    Debug.Print FileName
    Open App.path + "\" + FileName For Binary Access Read Lock Read As #1
    ReDim LoadBytes(LOF(1) - 1)
    Get #1, , LoadBytes
    Close #1
End Function
Public Sub SaveBytes(ByVal FileName As String, ByRef Biteritos() As Byte)
   
    Open App.path + "\RESOURCE\MELKOR_PNGS\" + FileName For Binary Access Write Lock Write As #1
    Put #1, , Biteritos
    Close #1
    
End Sub
Private Sub SetAmbientColor(ByVal red As Byte, _
                            ByVal green As Byte, _
                            ByVal blue As Byte, _
                            ByVal Alpha As Byte)

    Dim Uniform As wGL_Uniform
    
    Uniform.X = red / 255: Uniform.Y = green / 255: Uniform.Z = blue / 255: Uniform.W = Alpha / 255
    
    AmbientColor.A = Alpha
    AmbientColor.r = red
    AmbientColor.g = green
    AmbientColor.b = blue
    
    Call wGL_Graphic.Use_Uniform(&H0, False, Uniform, 1)
    
End Sub

'Sets a Grh animation to loop indefinitely.

Public Function GetElapsedTime() As Single
    
    On Error GoTo GetElapsedTime_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time    As Currency
    Static end_time   As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then Call QueryPerformanceFrequency(timer_freq)
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

    
    Exit Function

GetElapsedTime_Err:
    Resume Next
End Function
Public Function GetTickCount() As Long
    ' Devolvemos el valor absoluto de la cantidad de ticks que paso desde que prendimos la PC
    
    GetTickCount = (timeGetTime And &H7FFFFFFF)
    
End Function
