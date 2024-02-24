Attribute VB_Name = "mDamages"
Option Explicit
 
Const DAMAGE_TIME   As Integer = 57

Const DAMAGE_FONT_S As Byte = 14
 
Enum EDType

    d_DañoUser = 1
    d_DañoUserSpell = 2
    d_DañoNpc = 3
    d_DañoNpcSpell = 4
    d_CurarSpell = 5
    d_Apuñalar = 6
    d_DiamontRed = 7
    d_DiamontBlue = 8
    d_AddMan = 9
    d_AddExp = 10
    d_AddExpBonus = 11
    d_AddGld = 12
    d_AddGldBonus = 13
    d_Aniquilado = 14
    d_AniquiladoPor = 15
    d_AddMagicWord = 16
    d_DañoNpc_Critico = 17
    d_Fallas = 18
End Enum
 
Type DList

    DamageVal      As Long  'Cantidad de daño.
    ColorRGB       As Long     'Color.
    DamageType     As EDType   'Tipo, se usa para saber si es apu o no.
    TimeRendered   As Integer  'Tiempo transcurrido.
    Downloading    As Byte     'Contador para la posicion Y.
    Activated      As Boolean  'Si está activado..
    TimeDesvanecimiento As Byte
     
    Intervalo As Long
    OptionalText As String

End Type

Public BytesEnviados As Long
Public BytesTotal As Long

Sub CreateDamage(ByVal X As Byte, _
                 ByVal Y As Byte, _
                 ByVal DamageValue As Long, _
                 ByVal edMode As Byte, _
                 ByVal Texto As String)

    mDamages.ClearDamage X, Y
        
    'g_Swarm.Insert 7, edMode, X, Y, 2, 2
    
    With MapData(X, Y).Damage
    
        .Activated = True
        .DamageType = edMode
        .DamageVal = DamageValue
        .TimeRendered = 0
        .Downloading = 0
        .TimeDesvanecimiento = 255
        .Intervalo = FrameTime

        
        Select Case .DamageType

            Case EDType.d_AddExp
                .OptionalText = "Exp +"

            Case EDType.d_AddExpBonus
                .OptionalText = "ExpBonus +"

            Case EDType.d_AddGld
                .OptionalText = "Oro +"

            Case EDType.d_AddGldBonus
                .OptionalText = "OroBonus +"
            
            Case EDType.d_Aniquilado
                .OptionalText = "¡Aniquilado!"
            
            Case EDType.d_DañoNpc_Critico
                .OptionalText = "Critico +"
                
            Case EDType.d_Fallas
                .OptionalText = "Fallas"
                
            Case Else
                .OptionalText = vbNullString
            
        End Select
        
        If .OptionalText = vbNullString Then
            .OptionalText = Texto
        End If
    End With
       
End Sub

Sub DrawDamage(ByVal X As Byte, _
               ByVal Y As Byte, _
               ByVal PixelX As Integer, _
               ByVal PixelY As Integer)
       
    ' @ Dibuja un daño
    
    With MapData(X, Y).Damage
        
        If (FrameTime - .Intervalo) >= 2000 Then
            ClearDamage X, Y
            Exit Sub
        End If
        
        If (.DamageVal > 0 Or .DamageVal = -1 Or .DamageVal = -2) Then
            
            If .TimeRendered < DAMAGE_TIME Then
                'Sumo el contador del tiempo.
                If FrameTime - .Intervalo > 15 Then
                    .TimeRendered = .TimeRendered + 1
                    .Intervalo = FrameTime
                End If

                If (.TimeRendered * 0.5) > 0 Then
                    .Downloading = (.TimeRendered * 0.5)
                End If
                    
                .ColorRGB = ModifyColour(255, .DamageType)

                If .DamageVal = -1 Then
                    Mod_TileEngine.Draw_Text f_Tahoma, DAMAGE_FONT_S, PixelX, PixelY - .Downloading, 0#, 0#, .ColorRGB, FONT_ALIGNMENT_BASELINE, .OptionalText, True
                ElseIf .DamageVal = -2 Then
                    Mod_TileEngine.Draw_Text f_Medieval, 18, PixelX, PixelY - .Downloading, 0#, 0#, .ColorRGB, FONT_ALIGNMENT_BASELINE Or FONT_ALIGNMENT_CENTER, .OptionalText, False
                Else
                    If .DamageType = d_AniquiladoPor Then
                        Mod_TileEngine.Draw_Text f_Tahoma, DAMAGE_FONT_S, PixelX, PixelY - .Downloading, 0#, 0#, .ColorRGB, FONT_ALIGNMENT_BASELINE, "¡" & CharList(.DamageVal).Nombre & " Fight!", True
                    Else
                        Mod_TileEngine.Draw_Text f_Tahoma, DAMAGE_FONT_S, PixelX, PixelY - .Downloading, 0#, 0#, .ColorRGB, FONT_ALIGNMENT_BASELINE, .OptionalText & " " & .DamageVal, True
                    End If
                    
                End If
     
                'Si llego al tiempo lo limpio
                If .TimeRendered >= DAMAGE_TIME Then
                    ClearDamage X, Y
                End If
            End If
        End If
             
    End With
  
End Sub
 
Sub ClearDamage(ByVal X As Byte, ByVal Y As Byte)
    
    Dim NullDamage As DList
    MapData(X, Y).Damage = NullDamage

End Sub
 
Function ModifyColour(ByVal Alpha As Byte, ByVal DamageType As Byte) As Long
    
    Select Case DamageType
             
        Case EDType.d_Apuñalar
            ModifyColour = ARGB(255, 255, 184, Alpha)
                         
        Case EDType.d_DañoUser
            ModifyColour = ARGB(255, 40, 40, Alpha)
                
        Case EDType.d_CurarSpell
            ModifyColour = ARGB(26, 213, 13, Alpha)
              
        Case EDType.d_DañoNpc
            ModifyColour = ARGB(255, 40, 40, Alpha)
              
        Case EDType.d_DañoNpcSpell
            ModifyColour = ARGB(255, 40, 40, Alpha)
              
        Case EDType.d_DañoUserSpell
            ModifyColour = ARGB(200, 255, 230, Alpha)

        Case EDType.d_DiamontRed
            ModifyColour = ARGB(254, 227, 0, Alpha)
              
        Case EDType.d_DiamontBlue
            ModifyColour = ARGB(254, 227, 0, Alpha)
            
        Case EDType.d_AddMan
            ModifyColour = ARGB(70, 255, 200, Alpha)
            
        Case EDType.d_AddExp
            ModifyColour = ARGB(255, 215, 0, Alpha)
        
        Case EDType.d_AddExpBonus
            ModifyColour = ARGB(70, 255, 0, Alpha)
        
        Case EDType.d_AddGld
            ModifyColour = ARGB(255, 255, 0, Alpha)
        
        Case EDType.d_AddGldBonus
            ModifyColour = ARGB(255, 255, 0, Alpha)
            
        Case EDType.d_Aniquilado
            ModifyColour = ARGB(255, 100, 50, Alpha)
            
        Case EDType.d_AniquiladoPor
            ModifyColour = ARGB(255, 0, 0, Alpha)
            
        Case EDType.d_AddMagicWord
            ModifyColour = ARGB(140, 240, 220, Alpha)
            
        Case EDType.d_DañoNpc_Critico
            ModifyColour = ARGB(50, 240, 165, Alpha)
            
        Case EDType.d_Fallas
            ModifyColour = ARGB(220, 220, 220, Alpha)
    End Select
       
End Function

