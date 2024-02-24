Attribute VB_Name = "AI_NPCS"
Option Explicit

' Inteligencia creada por Argentum Game Community

'Public Enum eBot_Action
'ACTION_AFK = 0 ' La criatura está quieta. Esperando una orden
'ACTION_MOVEMENT = 1 'La criatura está yendo del Punto A => Punto B
'ACTION_RETURN = 2 ' 'La Criatura está yendo del Punto B => Punto A
'End Enum

Public Sub GreedyWalkTo(ByVal NpcIndex As Integer, _
                        ByVal TargetMap As Integer, _
                        ByVal TargetX As Integer, _
                        ByVal TargetY As Integer)

    '***************************************************
    'Author: Unknown
    '  Este procedimiento es llamado cada vez que un NPC deba ir
    '  a otro lugar en el mismo TargetMap. Utiliza una técnica
    '  de programación greedy no determinística.
    '  Cada paso azaroso que me acerque al destino, es un buen paso.
    '  Si no hay mejor paso válido, entonces hay que volver atrás y reintentar.
    '  Si no puedo moverme, me considero piketeado
    '  La funcion es larga, pero es O(1) - orden algorítmico temporal constante
    'Last Modification: 26/09/2010
    'Rapsodius - Changed Mod by And for speed
    '26/09/2010: ZaMa - Now make movements as normal npcs, which allows to kick caspers and invisible admins.
    '***************************************************
    On Error GoTo ErrHandler

    Dim NpcX      As Integer

    Dim NpcY      As Integer

    Dim RandomDir As Integer
    
    With Npclist(NpcIndex).Pos

        If .Map <> TargetMap Then Exit Sub
        
        NpcX = .X
        NpcY = .Y

    End With
    
    ' Arrived to destination
    If (NpcX = TargetX And NpcY = TargetY) Then Exit Sub
    
    ' Try to reach target
    If (NpcX > TargetX) Then
    
        ' Target is Down-Left
        If (NpcY < TargetY) Then
            
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move down
            
                If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                End If
                
            Else

                ''random first move
                If MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                End If
                
            End If
            
            ' Target is Up-Left
        ElseIf (NpcY > TargetY) Then
        
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move up
            
                If MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                End If
                
            Else

                ''random first move
                If MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                End If
                
            End If
            
            ' Target is Straight Left
        Else
        
            If MoveNPCChar(NpcIndex, eHeading.WEST) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                Exit Sub

            Else

                ' If moves to east, algorithm fails because starts a loop
                'Call MoveFailed(NpcIndex)
            End If
            
        End If
    
    ElseIf (NpcX < TargetX) Then
        
        ' Target is Down-Right
        If (NpcY < TargetY) Then
            
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move down
            
                If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                End If
                
            Else    ''random first move
                
                If MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                End If
                
            End If
        
            ' Target is Up-Right
        ElseIf (NpcY > TargetY) Then
        
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move up
            
                If MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                End If
                
            Else
            
                If MoveNPCChar(NpcIndex, eHeading.EAST) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                    Exit Sub

                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                    Exit Sub

                End If
                
            End If
        
            ' Target is Straight Right
        Else
        
            If MoveNPCChar(NpcIndex, eHeading.EAST) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                Exit Sub

            Else

                ' If moves to west, algorithm fails because starts a loop
                'Call MoveFailed(NpcIndex)
            End If
            
        End If
    
        ' Target is straight Up/Down
    Else
    
        ' Target is straight Up
        If (NpcY > TargetY) Then
        
            If MoveNPCChar(NpcIndex, eHeading.NORTH) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                Exit Sub

            Else

                ' If moves to south, algorithm fails because starts a loop
                'Call MoveFailed(NpcIndex)
            End If
            
            ' Target is straight Down
        Else
        
            If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then

                Exit Sub

            ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then

                Exit Sub

            Else

                ' If moves to north, algorithm fails because starts a loop
                'Call MoveFailed(NpcIndex)
            End If
            
        End If
        
    End If
    
    Exit Sub

ErrHandler:
    LogError ("Error en GreedyWalkTo. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Function AI_BestTarget(ByVal NpcIndex As Integer) As Integer
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer
        
        Dim mapa               As Integer

        Dim NPCPosX            As Integer

        Dim NPCPosY            As Integer
        
        Dim UserIndex          As Integer

        Dim Counter            As Long
        
        Dim BestTargetDistance As Integer

        Dim Distance           As Integer
        
        With Npclist(NpcIndex).Pos
            mapa = .Map
            NPCPosX = .X
            NPCPosY = .Y
        End With
        
        Dim CounterStart As Long

        Dim CounterEnd   As Long

        Dim CounterStep  As Long
        
             Dim query() As Collision.UUID
        Call ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

        
        ' To avoid that all attack the same target
        CounterStep = RandomNumber(0, 1)

       If CounterStep = 1 Then
            CounterStart = 1
            CounterEnd = UBound(query)
        Else
            CounterStart = UBound(query)
            CounterEnd = 1
            CounterStep = -1
        End If
        
        ' Search for the best user target
        For Counter = CounterStart To CounterEnd Step CounterStep
        
             UserIndex = query(Counter).Name
            
  

                ' Can be atacked? If it's blinded, doesn't count.
                If UserAtacable(UserIndex, NpcIndex) And UserList(UserIndex).flags.Ceguera = 0 Then

                    ' if previous user choosen, select the better
                    If BestTarget <> 0 Then
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NPCPosX, NPCPosY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance
                        End If

                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NPCPosX, NPCPosY)
                    End If
              
            End If
                
        Next Counter

    AI_BestTarget = BestTarget

    Exit Function

ErrHandler:
    LogError ("Error en KingBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function
Private Function UserDistance(ByVal UserIndex As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 24/06/2010
        'Calculates the user distance according to the given coords.
        '***************************************************
        '<EhHeader>
        On Error GoTo UserDistance_Err
        '</EhHeader>

100     With UserList(UserIndex)
102         UserDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)
        End With
    
        '<EhFooter>
        Exit Function

UserDistance_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.AI_NPCS.UserDistance " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function UserAtacable(ByVal UserIndex As Integer, _
                              ByVal NpcIndex As Integer, _
                              Optional ByVal CheckVisibility As Boolean = True, _
                              Optional ByVal AttackAdmin As Boolean = True) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/10/2010
        'DEtermines whether the user can be atacked or not
        '05/10/2010: ZaMa - Now doesn't allow to attack admins sometimes.
        '***************************************************
        '<EhHeader>
        On Error GoTo UserAtacable_Err
        '</EhHeader>

100     With UserList(UserIndex).flags
102         UserAtacable = Not .EnConsulta And .AdminInvisible = 0 And .AdminPerseguible And .Muerto = 0
                       
104         If CheckVisibility Then
106             UserAtacable = UserAtacable And .Oculto = 0 And .Invisible = 0
            End If
        
108         If Not AttackAdmin Then
110             UserAtacable = UserAtacable And (Not EsGm(UserIndex))
            End If

        End With
                        
        '<EhFooter>
        Exit Function

UserAtacable_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.AI_NPCS.UserAtacable " & _
           "at line " & Erl
    
    '</EhFooter>
End Function

