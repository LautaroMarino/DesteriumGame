Attribute VB_Name = "mMaps"
' Modulo dedicado a todos los sonidos, efectos y cosas de sonidos que se haran. CARGADO DE SONIDOS AMBIENTALES.

Option Explicit


Public Type tMapInfo_Sounds
    X As Byte
    Y As Byte
    Sound As Integer
        
End Type

Public Type tMapConfig
    Name As String
    
    Sound As Integer
    Sounds() As tMapInfo_Sounds
    
    
    RenderX As Integer ' PosX donde dibuja el centro del mapa (FrmMapa)
    RenderY As Integer ' PosY donde dibuja el centro del mapa (FrmMapa)
End Type


Public LastMap As Integer
Public MapConfig() As tMapConfig

Sub ILoad_MapInfo()

    Dim Manager As clsIniManager
    Dim A As Long, b As Long
    Dim Text As String
    
    Set Manager = New clsIniManager
    
    Manager.Initialize IniPath & "Mapinfo.ind"
    
    LastMap = Val(Manager.GetValue("INIT", "LAST"))
    
    ReDim MapConfig(1 To LastMap) As tMapConfig
    
    For A = 1 To LastMap
        With MapConfig(A)
            .Name = Manager.GetValue(CStr(A), "NAME")
            
            .Sound = Val(Manager.GetValue(CStr(A), "SOUND"))
            
            ReDim .Sounds(0 To .Sound) As tMapInfo_Sounds
            
            For b = 1 To .Sound
                Text = Manager.GetValue(CStr(A), "S" & b)
                .Sounds(b).Sound = Val(ReadField(1, Text, 45))
                .Sounds(b).X = Val(ReadField(2, Text, 45))
                .Sounds(b).Y = Val(ReadField(3, Text, 45))
            Next b
            
            
            .RenderX = Val(Manager.GetValue(CStr(A), "RenderX"))
            .RenderY = Val(Manager.GetValue(CStr(A), "RenderY"))
            
        End With
    Next A
    
    Set Manager = Nothing
End Sub

' @ Carga los sonidos de un mapa
Sub IMapInitial_Sound(ByVal UserMap As Integer)
    
    Dim A As Long
    Dim PathAmbient As String
    
    If UserMap > LBound(MapConfig) Then Exit Sub
        
    PathAmbient = "AMBIENT\"
    
    With MapConfig(UserMap)

        For A = 1 To .Sound

            With MapData(.Sounds(A).X, .Sounds(A).Y)

                If .SoundSource > 0 Then
                    Call Audio.DeleteSource(.SoundSource, True)
                End If
                 
                  .SoundSource = Audio.CreateSource(MapConfig(UserMap).Sounds(A).X, MapConfig(UserMap).Sounds(A).Y)
                  Call Audio.PlayEffect(PathAmbient & CStr(MapConfig(UserMap).Sounds(A).Sound) & ".wav", .SoundSource, True)
                  'Call Audio.UpdateSource(MapConfig(UserMap).Sounds(A).Sound, MapConfig(UserMap).Sounds(A).X, MapConfig(UserMap).Sounds(A).Y)
            End With
            
        Next A
    End With
        
End Sub
