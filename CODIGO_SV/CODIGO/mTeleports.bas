Attribute VB_Name = "mTeleports"
Option Explicit


Public Const TELEPORTS_DELAY_INVOKER As Long = 5000

Public Type tTeleportsCounters
        Duration As Long
        Invocation As Long
        
End Type

Public Type tTeleports
    Active As Boolean
    ObjIndex As Integer
    UserIndex As Integer
    
    PositionInvoker As WorldPos         ' Posición donde comienza a crear el teleports, se pueden posicionar npcs y usuarios, y eso hace que altere la pos del warpeo final.
    Position  As WorldPos                   ' Posición donde aparece el Teleport.
    PositionWarp As WorldPos            ' Posición donde te lleva el Teleport.
    Counters As tTeleportsCounters
    
    TeleportObj As Integer                 ' Teleport objeto que va a utilizar.
    FxInvoker As Integer                    ' Animación mientras se crea el Teleport

    CanGuild As Boolean
    CanParty As Boolean
End Type


Public Const TELEPORT_MAX_SPAWN As Byte = 100       ' Máximo de Teleports que hay en el mundo.
Public Teleports(1 To TELEPORT_MAX_SPAWN) As tTeleports


' @ Busca un slot libre para poder crear el teleport
Private Function Teleports_FreeSlot() As Integer
        '<EhHeader>
        On Error GoTo Teleports_FreeSlot_Err
        '</EhHeader>
        Dim A As Long
    
100     For A = 1 To TELEPORT_MAX_SPAWN
102         With Teleports(A)
104             If .Active = False Then
106                 Teleports_FreeSlot = A
                    Exit Function
                End If
            End With
    
108     Next A
        '<EhFooter>
        Exit Function

Teleports_FreeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_FreeSlot " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub Teleports_Loop()
        '<EhHeader>
        On Error GoTo Teleports_Loop_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To TELEPORT_MAX_SPAWN

102         With Teleports(A)

104             If .Active Then

                    ' @ Tiempo que tarda el Teleport en aparecer en el mapa.
106                 If .Counters.Invocation > 0 Then
108                     .Counters.Invocation = .Counters.Invocation - 1

                        'Call SendData(SendTarget.ToPCArea, .UserIndex, PrepareMessageUpdateBar(UserList(.UserIndex).Char.CharIndex, eTypeBar.eTeleportInvoker, .Counters.Invocation, ObjData(.ObjIndex).TimeWarp))
110                      Call SendData(SendTarget.ToPCArea, .UserIndex, PrepareMessageUpdateBarTerrain(.Position.X, .Position.Y, eTypeBar.eTeleportInvoker, .Counters.Invocation, ObjData(.ObjIndex).TimeWarp))
                     
112                     If .Counters.Invocation = 0 Then
114                         Call Teleports_Spawn(A)
                        End If
                
                    Else

                        ' @ Duración del Teleport hasta que desaparece.
116                     If .Counters.Duration > 0 Then
118                         .Counters.Duration = .Counters.Duration - 1
                        
120                         If .Counters.Duration = 0 Then
122                             Call Teleports_Remove(A)

                            End If

                        End If

                    End If
            
                End If
        
            End With
    
124     Next A

        '<EhFooter>
        Exit Sub

Teleports_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_Loop " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Teleports_DeterminateSoundWarp(ByVal UserIndex As Integer, _
                                           ByVal ObjIndex As Integer, _
                                           ByVal SourceX As Integer, _
                                           ByVal SourceY As Integer)
        '<EhHeader>
        On Error GoTo Teleports_DeterminateSoundWarp_Err
        '</EhHeader>
    
        Dim Sound As Integer
    
100     With ObjData(ObjIndex)

102         Select Case .TimeWarp
        
                Case 11
104                 Sound = eSound.sWarp10s

106             Case 21
108                 Sound = eSound.sWarp20s

110             Case 31
112                 Sound = eSound.sWarp30s

114             Case 61
116                 Sound = eSound.sWarp60s

118             Case Else
                    Exit Sub
        
            End Select
        
120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Sound, SourceX, SourceY, 0, False, True))
        
        End With

        '<EhFooter>
        Exit Sub

Teleports_DeterminateSoundWarp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_DeterminateSoundWarp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Teleports_AddNew(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            ByVal Map As Integer, _
                            ByVal X As Byte, _
                            ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo Teleports_AddNew_Err
        '</EhHeader>
                                            
        Dim Slot As Integer

        Dim Time As Double

        Dim nPos As WorldPos
    
100     Time = GetTime
102     Slot = Teleports_CheckWarp(UserIndex, Map, X, Y, ObjIndex, Time)
    
104     If Slot > 0 Then

106         With Teleports(Slot)
            
108             nPos.Map = Map
110             nPos.X = X
112             nPos.Y = Y
114             ClosestStablePos nPos, nPos

116             If nPos.Map = 0 Or nPos.X = 0 Or nPos.Y = 0 Then Exit Sub
                  
118             .Active = True
120             .ObjIndex = ObjIndex
           
122             .Counters.Invocation = ObjData(ObjIndex).TimeWarp
124             .TeleportObj = ObjData(ObjIndex).TeleportObj

126             .Position.Map = nPos.Map
128             .Position.X = nPos.X
130             .Position.Y = nPos.Y
132             .PositionInvoker = .Position
            
134             .PositionWarp.Map = ObjData(ObjIndex).Position.Map
136             .PositionWarp.X = ObjData(ObjIndex).Position.X
138             .PositionWarp.Y = ObjData(ObjIndex).Position.Y
            
140             .UserIndex = UserIndex
            
142             .FxInvoker = ObjData(ObjIndex).FX
                
                Call Teleports_DeterminateSoundWarp(UserIndex, ObjIndex, .Position.X, .Position.Y)
                
144             With UserList(UserIndex)
146                 .flags.TeleportInvoker = Slot
148                 .flags.LastInvoker = GetTime
                    '  .Char.loops = INFINITE_LOOPS
                    ' .Char.FX = Teleports(Slot).FxInvoker
150                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, Teleports(Slot).FxInvoker, , , False))
154                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateBarTerrain(Teleports(Slot).Position.X, Teleports(Slot).Position.Y, eTypeBar.eTeleportInvoker, Teleports(Slot).Counters.Invocation, ObjData(ObjIndex).TimeWarp))
156                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(Teleports(Slot).Position.X, Teleports(Slot).Position.Y, ObjData(Teleports(Slot).TeleportObj).FX, -1))

                End With
        
            End With

        End If

        '<EhFooter>
        Exit Sub

Teleports_AddNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_AddNew " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' @ Comprueba que el usuarip ueda teletransportarse
Private Function Teleports_CheckWarp(ByVal UserIndex As Integer, _
                                     ByVal Map As Integer, _
                                     ByVal X As Byte, _
                                     ByVal Y As Byte, _
                                     ByVal ObjIndex As Integer, _
                                     ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Teleports_CheckWarp_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

102         If Not InMapBounds(Map, X, Y) Then
                Exit Function
            End If
              
              If .flags.Meditando Then Exit Function
              If ObjData(ObjIndex).OBJType <> otTeleportInvoker Then Exit Function ' @ Seleccionó otro objeto despues del teleport.
104         If .Pos.Map = ObjData(ObjIndex).Position.Map Then Exit Function
106         If .flags.TeleportInvoker > 0 Then Exit Function ' @Esta invocando otro
108         If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then Exit Function ' @ Hay un objeto.
110         If MapData(Map, X, Y).TileExit.Map > 0 Then Exit Function     ' @ Hay otro traslado
112         If MapData(Map, X, Y).NpcIndex > 0 Then Exit Function ' @ Hay una criatura
114         If MapData(Map, X, Y).UserIndex > 0 Then Exit Function ' @ Hay un usuario
116         If MapData(Map, X, Y).Blocked > 0 Then Exit Function ' @ Está bloqueado
118         If MapData(Map, X, Y).trigger > 0 Then Exit Function  ' @ Hay Trigger
            If MapData(Map, X, Y).TeleportIndex > 0 Then Exit Function  ' @ Hay otro Portal!
            
            ' Es un teleport inalcanzable
120         If MapData(Map, X - 1, Y).Blocked > 0 And MapData(Map, X + 1, Y).Blocked > 0 And MapData(Map, X, Y + 1).Blocked > 0 And MapData(Map, X, Y - 1).Blocked > 0 Then Exit Function ' @ Está bloqueado
        
122         If ObjData(ObjIndex).PuedeInsegura = 0 And MapInfo(.Pos.Map).Pk Then
124             Call WriteConsoleMsg(UserIndex, "¡Este teleport no puede ser usado desde zona insegura.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If
        
126         If ObjData(ObjIndex).PuedeInsegura = 1 And MapInfo(.Pos.Map).Pk = False Then
128             Call WriteConsoleMsg(UserIndex, "¡Este teleport solo puede ser usado desde zona insegura!", FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If
        
130         If ObjData(ObjIndex).LvlMin > .Stats.Elv Then
132             Call WriteConsoleMsg(UserIndex, "Debes ser Nivel " & ObjData(ObjIndex).LvlMin & " para poder invocar el Portal.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If
        
134         If ObjData(ObjIndex).LvlMax < .Stats.Elv Then
136             Call WriteConsoleMsg(UserIndex, "El portal puede ser invocado por personas inferiores la nivel " & ObjData(ObjIndex).LvlMax, FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If
        
138         If (GetTime - UserList(UserIndex).flags.LastInvoker) <= TELEPORTS_DELAY_INVOKER Then
140             Call WriteConsoleMsg(UserIndex, "¡Debes esperar algunos segundos antes de volver a invocar un Portal!", FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If
            
            If ObjData(ObjIndex).Dead = 1 And .flags.Muerto = 0 Then
                Call WriteConsoleMsg(UserIndex, "¡Este portal solo puede ser invocado estando muerto!", FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If

142         If .flags.SlotEvent > 0 Then Exit Function
144         If .flags.SlotFast > 0 Then Exit Function
146         If .flags.Desafiando > 0 Then Exit Function
148         If .Counters.Pena > 0 Then Exit Function
            
            If ObjData(ObjIndex).LvlMin >= 25 Then
                If Not .Stats.UserSkills(eSkill.Navegacion) >= 35 Then
                    Call WriteConsoleMsg(UserIndex, "¡Debes tener al menos una barca para poder viajar! Recuerda además tener la capacidad de usar la embarcación según tus skills.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Function
                End If
                
                If Not TieneObjetos(474, 1, UserIndex) And Not TieneObjetos(475, 1, UserIndex) And Not TieneObjetos(476, 1, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "¡Debes tener al menos una barca para poder viajar! Recuerda además tener la capacidad de usar la embarcación según tus skills.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Function
                End If
            End If
            
150         Teleports_CheckWarp = Teleports_FreeSlot

        End With

        '<EhFooter>
        Exit Function

Teleports_CheckWarp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_CheckWarp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

' @ El teleport aparece en el mapa
Private Sub Teleports_Spawn(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo Teleports_Spawn_Err
        '</EhHeader>
    
        Dim Position    As WorldPos

        Dim nPos        As WorldPos

        Dim ObjTeleport As Obj
    
100     With Teleports(Slot)
102         If Not TieneObjetos(.ObjIndex, 1, .UserIndex) Then
104             Call Teleports_Remove(Slot)
                Exit Sub
            End If

108         If .PositionWarp.X = 0 And .PositionWarp.Y = 0 And .PositionWarp.Map > 0 Then
110             .PositionWarp.X = RandomNumber(20, 80)
112             .PositionWarp.Y = RandomNumber(20, 80)
        
114         ElseIf .PositionWarp.X = 0 And .PositionWarp.Y = 0 And .PositionWarp.Map = 0 Then
116             .PositionWarp.Map = UserList(.UserIndex).Hogar
118             .PositionWarp.X = RandomNumber(20, 80)
120             .PositionWarp.Y = RandomNumber(20, 80)
            End If
        
            ' @ Teleport que invoco en mi mapa
122         ClosestStablePos .Position, nPos
              
            If nPos.Map = 0 Or nPos.X = 0 Or nPos.Y = 0 Then
                Call Teleports_Remove(Slot)
                Exit Sub
            End If
            
              .Counters.Duration = ObjData(.ObjIndex).TimeDuration
            
124         MapData(nPos.Map, nPos.X, nPos.Y).TileExit = .PositionWarp
126         .Position = nPos
        
128         ObjTeleport.ObjIndex = .TeleportObj
130         ObjTeleport.Amount = 1
        
132         Call MakeObj(ObjTeleport, nPos.Map, nPos.X, nPos.Y)
134         MapData(nPos.Map, nPos.X, nPos.Y).TeleportIndex = Slot
        
            'Quitamos del inv el item
136         If ObjData(.ObjIndex).RemoveObj > 0 Then
138             Call QuitarObjetos(.ObjIndex, ObjData(.ObjIndex).RemoveObj, .UserIndex)
            End If
        
140         Call Teleports_Reset_Effect(Slot, .UserIndex)
            'Call SendData(SendTarget.ToPCArea, .UserIndex, PrepareMessageStopWaveMap(.Position.X, .Position.Y, False))
        End With

        '<EhFooter>
        Exit Sub

Teleports_Spawn_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_Spawn " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Teleports_Remove(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo Teleports_Remove_Err
        '</EhHeader>

        Dim TeleportNull As tTeleports
    
100     With Teleports(Slot)
        
102         With MapData(.Position.Map, .Position.X, .Position.Y)
104             .TeleportIndex = 0
            
106             If .ObjInfo.ObjIndex > 0 Then
108                 Call EraseObj(1, Teleports(Slot).Position.Map, Teleports(Slot).Position.X, Teleports(Slot).Position.Y)

                End If
            
110             If .TileExit.Map > 0 Then
112                 .TileExit.Map = 0
114                 .TileExit.X = 0
116                 .TileExit.Y = 0

                End If

            End With
        
118         UserList(.UserIndex).flags.TeleportInvoker = 0
             ' UserList(.UserIndex).flags.LastInvoker = 0
120         UserList(.UserIndex).Char.FX = 0
122         UserList(.UserIndex).Char.loops = 0
        
        
            
124         Call Teleports_Reset_Effect(Slot, .UserIndex)
        
        End With

126     Teleports(Slot) = TeleportNull
    
        '<EhFooter>
        Exit Sub

Teleports_Remove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_Remove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Teleports_Cancel(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Teleports_Cancel_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

102         If .flags.TeleportInvoker = 0 Then Exit Sub
        
104         If Teleports(.flags.TeleportInvoker).Counters.Duration > 0 Then
106             If Teleports(.flags.TeleportInvoker).Counters.Invocation = 0 Then Exit Sub
            End If
        
               ' Forzamos a parar el sonido
108         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageStopWaveMap(.Pos.X, .Pos.Y, True))
110         Call Teleports_Reset_Effect(.flags.TeleportInvoker, UserIndex)
112         Call Teleports_Remove(.flags.TeleportInvoker)
        
     
        End With

        '<EhFooter>
        Exit Sub

Teleports_Cancel_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_Cancel " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Teleports_Reset_Effect(ByVal Slot As Byte, ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Teleports_Reset_Effect_Err
        '</EhHeader>

        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateBar(UserList(UserIndex).Char.CharIndex, eTypeBar.eTeleportInvoker, 0, 0))
    
100     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateBarTerrain(Teleports(Slot).Position.X, Teleports(Slot).Position.Y, eTypeBar.eTeleportInvoker, 0, 0))
102     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0, , , False))
104     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(Teleports(Slot).PositionInvoker.X, Teleports(Slot).PositionInvoker.Y, 0, 0))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(0, Teleports(Slot).Position.X, Teleports(Slot).Position.Y, 0, False, True))
        
        '<EhFooter>
        Exit Sub

Teleports_Reset_Effect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mTeleports.Teleports_Reset_Effect " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
