Attribute VB_Name = "AI"
'Argentum Online 0.12.2
'Copyright (C) 2002 Mï¿½rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 nï¿½mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Cï¿½digo Postal 1900
'Pablo Ignacio Mï¿½rquez

Option Explicit



Public Enum TipoAI

    Estatico = 1
    MueveAlAzar = 3
    NpcDefensa = 4
    SigueAmo = 8
    eNpcAtacaNpc = 9
    GuardiaPersigueNpc = 10
    
    NpcDagaRusa = 11
    NpcGranBestia = 12
    
    ArghalSacerdote = 13
    IntelligenceMax = 14            ' @ Inteligencia aplicada para las criaturas inteligentes
    Caminata = 15
    Invasion = 16
End Enum

'Damos a los NPCs el mismo rango de visiï¿½n que un PJ


Public Const RANGO_VISION_x As Byte = 8
Public Const RANGO_VISION_y As Byte = 6



Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo RestoreOldMovement_Err
        '</EhHeader>

100     With Npclist(NpcIndex)

102         If .MaestroUser = 0 Then
104             .Movement = .flags.OldMovement
106             .Hostile = .flags.OldHostil
108             .flags.AttackedBy = vbNullString
110             .flags.KeepHeading = 0

            End If

        End With

        '<EhFooter>
        Exit Sub

RestoreOldMovement_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.AI.RestoreOldMovement " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo NpcLanzaUnSpell_Err

        '</EhHeader>

        '**************************************************************
        'Author: Unknown
        'Last Modify by: -
        'Last Modify Date: -
        '**************************************************************
100     With UserList(UserIndex)

102         If .flags.Invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
              If Not Intervalo_CriatureAttack(NpcIndex, False) Then Exit Sub
        End With
    
        Dim K As Integer

104     K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
106     Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(K))
        '<EhFooter>
        Exit Sub

NpcLanzaUnSpell_Err:
        LogError Err.description & vbCrLf & "in NpcLanzaUnSpell " & "at line " & Erl

        '</EhFooter>
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo NpcLanzaUnSpellSobreNpc_Err

        '</EhHeader>

        Dim K As Integer

100     K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
102     Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(K), True)
        '<EhFooter>
        Exit Sub

NpcLanzaUnSpellSobreNpc_Err:
        LogError Err.description & vbCrLf & "in NpcLanzaUnSpellSobreNpc " & Npclist(NpcIndex).Name & ". " & "at line " & Erl

        '</EhFooter>
End Sub


Public Sub Events_AI_DagaRusa(ByVal NpcIndex As Integer, _
                            Optional ByVal Init As Boolean = False)


        On Error GoTo Events_AI_DagaRusa_Err


        Dim UserIndex  As Integer

        Dim Npc        As Npc

        Dim LoopC      As Integer

        Dim SlotEvent  As Integer

        Dim tHeading   As eHeading

        Dim Pos        As WorldPos
          
        Static Pasaron As Byte
          

100     Npc = Npclist(NpcIndex)
102     SlotEvent = Npc.flags.SlotEvent
          
104     If Init Then
106         Pasaron = 0

            Exit Sub

        End If
        
108     With Events(SlotEvent)
              
            ' El NPC completa la ronda.
110         If Pasaron >= Npclist(NpcIndex).flags.InscribedPrevio Then
                
                DataRusa_SummonUser SlotEvent, False ' @ Vuelve a summonear a los usuarios
112             DagaRusa_ResetRonda SlotEvent
114             UserIndex = DagaRusa_NextUser(SlotEvent)
                  
116             Pos.Map = Npclist(NpcIndex).Pos.Map
118             Pos.X = UserList(UserIndex).Pos.X
120             Pos.Y = UserList(UserIndex).Pos.Y - 1
122             tHeading = FindDirection(Npclist(NpcIndex).Pos, Pos)
124             Call MoveNPCChar(NpcIndex, tHeading)
                  
126             If Npclist(NpcIndex).Pos.X = Pos.X Then
128                 Pasaron = 0
130                 Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed

                End If
                  
                Exit Sub

            End If
                      
132         UserIndex = DagaRusa_NextUser(SlotEvent)
              
134         If UserIndex > 0 Then
              
136             If Not (Distancia(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos) <= 1) Then
138                 Pos.Map = UserList(UserIndex).Pos.Map
140                 Pos.X = UserList(UserIndex).Pos.X
142                 Pos.Y = UserList(UserIndex).Pos.Y - 1
                                  
144                 tHeading = FindDirection(Npclist(NpcIndex).Pos, Pos)
146                 Call MoveNPCChar(NpcIndex, tHeading)
148                 Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, tHeading)
                Else
150                 Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, eHeading.SOUTH)

                End If
                  
152             If (Distancia(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos) <= 1) Then
154                 Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, SOUTH)
156                 .Users(UserList(UserIndex).flags.SlotUserEvent).Value = 1
                
158                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_IMPACTO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
160                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, FXSANGRE, 0))
                
162                 If RandomNumber(1, 100) <= .Prob Then
164                     Call UserDie(UserIndex)
                    Else
166                     Call WriteConsoleMsg(UserIndex, "¡Te has salvado! Pero no tendrás tanta suerte la próxima...", FontTypeNames.FONTTYPE_INFOGREEN)

                    End If

168                 Npclist(NpcIndex).Target = UserIndex
170                 Pasaron = Pasaron + 1

                End If
              
            End If
              
        End With


        Exit Sub

Events_AI_DagaRusa_Err:
    LogError Err.description & vbCrLf & "in Events_AI_DagaRusa " & "at line " & Erl

End Sub

Private Sub GeneralPathFinder(ByVal NpcIndex As Integer, _
                              ByVal TargetIndex As Integer, _
                              Optional ByVal TargetType As Byte = 0, _
                              Optional ByVal IsPet As Boolean = False)

    '---------------------------------------------------------------------------------------
    ' Module    : AI
    ' Author    : Anagrama
    ' Date      : 17/07/2016
    ' Purpose   : Busca el mejor camino hacia el usuario sobre la marcha.
    ' Aclaro que la lectura de esto es una tortura porque son puros if con el
    ' fin de ahorrar la mayor cantidad de recursos posibles.
    ' Todavia se puede mejorar muchisimo esto para optimizarlo mas.
    ' 07/10/2016: Anagrama - Ahora funciona contra NPCs.
    ' 17/11/2016: Anagrama - Modificado para funcionar bien con npcs de agua y tomar como valor si es una mascota o no.
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
  
    Dim AxisX     As Byte

    Dim AxisY     As Byte

    Dim MyHeading As Byte
    
    If TargetType = 0 Then
        If Distancia(Npclist(NpcIndex).Pos, UserList(TargetIndex).Pos) = 1 Then Exit Sub
    Else

        If Distancia(Npclist(NpcIndex).Pos, Npclist(TargetIndex).Pos) = 1 Then Exit Sub

    End If
    
    With Npclist(NpcIndex)

        If TargetType = 0 Then 'El target es un user.

            'Aca se encuentra la posicion relativa en los ejes X e Y
            If .Pos.Y < UserList(TargetIndex).Pos.Y Then
                AxisY = 0 'Npc arriba
            ElseIf .Pos.Y > UserList(TargetIndex).Pos.Y Then
                AxisY = 1 'Npc abajo
            Else
                AxisY = 2 'Npc vertical

            End If
            
            If .Pos.X > UserList(TargetIndex).Pos.X Then
                AxisX = 0 'Npc derecha
            ElseIf .Pos.X < UserList(TargetIndex).Pos.X Then
                AxisX = 1 'Npc izquierda
            Else
                AxisX = 2 'Npc horizontal

            End If

        Else 'El target es un npc.

            'Aca se encuentra la posicion relativa en los ejes X e Y
            If .Pos.Y < Npclist(TargetIndex).Pos.Y Then
                AxisY = 0 'Npc arriba
            ElseIf .Pos.Y > Npclist(TargetIndex).Pos.Y Then
                AxisY = 1 'Npc abajo
            Else
                AxisY = 2 'Npc vertical

            End If
            
            If .Pos.X > Npclist(TargetIndex).Pos.X Then
                AxisX = 0 'Npc derecha
            ElseIf .Pos.X < Npclist(TargetIndex).Pos.X Then
                AxisX = 1 'Npc izquierda
            Else
                AxisX = 2 'Npc horizontal

            End If

        End If
        
        'Asigna la direccion a la que revisar basandose en si es la primera vez que intenta buscar o no.
        If .flags.KeepHeading = 1 Then
            MyHeading = .Char.Heading
        Else

            If AxisY = 0 And AxisX <> 1 Then
                MyHeading = eHeading.SOUTH
            ElseIf AxisY <> 0 And AxisX = 0 Then
                MyHeading = eHeading.WEST
            ElseIf AxisY <> 1 And AxisX = 1 Then
                MyHeading = eHeading.EAST
            ElseIf AxisY = 1 And AxisX <> 0 Then
                MyHeading = eHeading.NORTH

            End If

            .flags.KeepHeading = 1

        End If
        
        'Segun la direccion de npc revisa su posicion en relacion a la pos del target,
        'seguido a eso intenta moverse eliminando la diferencia en el eje contrario
        'al que esta apuntando.
        Select Case MyHeading

            Case eHeading.NORTH

                If AxisX = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub

                            End If

                        End If

                    End If

                ElseIf AxisX = 0 Then

                    If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub

                            End If

                        End If

                    End If

                Else

                    If AxisY = 0 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub

                        End If

                    End If

                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub

                        End If

                    End If

                End If

            Case eHeading.WEST

                If AxisY = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub

                            End If

                        End If

                    End If

                ElseIf AxisY = 0 Then

                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub

                            End If

                        End If

                    End If

                Else

                    If AxisX = 0 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub

                        End If

                    End If
                    
                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub

                        End If

                    End If

                End If

            Case eHeading.SOUTH

                If AxisX = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub

                            End If

                        End If

                    End If

                ElseIf AxisX = 0 Then

                    If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub

                            End If

                        End If

                    End If

                Else

                    If AxisY = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub

                        End If

                    End If

                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub

                        End If

                    End If

                End If

            Case eHeading.EAST

                If AxisY = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub

                            End If

                        End If

                    End If

                ElseIf AxisY = 0 Then

                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    Else

                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub

                            End If

                        Else

                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub

                            End If

                        End If

                    End If

                Else

                    If AxisX = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub

                        End If

                    End If

                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub

                        End If

                    Else

                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub

                        End If

                    End If

                End If

        End Select

    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.number & "(" & Err.description & ") en Sub GeneralPathFinder de AI_NPC.bas")

End Sub
