Attribute VB_Name = "ModAI"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Public Const FUEGOFATUO      As Integer = 964
Public Const ELEMENTAL_VIENTO      As Integer = 963
Public Const ELEMENTAL_FUEGO      As Integer = 962

Public Const DIAMETRO_VISION_GUARDIAS_NPCS As Byte = 7

Public Sub NpcAI(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        'Debug.Print "NPC: " & NpcList(NpcIndex).Name
100     With Npclist(NpcIndex)
                
102         Select Case .Movement
                Case TipoAI.Estatico
                    ' Es un NPC estatico, no hace nada.
                    Exit Sub

104             Case TipoAI.MueveAlAzar
106                 If .Hostile = 1 Then
108                     Call PerseguirUsuarioCercano(NpcIndex)
                    Else
110                     Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
                    End If

112             Case TipoAI.NpcDefensa
114                 Call SeguirAgresor(NpcIndex)

116             Case TipoAI.eNpcAtacaNpc
118                 Call AI_NpcAtacaNpc(NpcIndex)

120             Case TipoAI.SigueAmo
122                 Call SeguirAmo(NpcIndex)

124             Case TipoAI.Caminata
126                 Call HacerCaminata(NpcIndex)

128             Case TipoAI.Invasion
130                 Call MovimientoInvasion(NpcIndex)

132             Case TipoAI.GuardiaPersigueNpc
134                 Call AI_GuardiaPersigueNpc(NpcIndex)

                Case TipoAI.NpcDagaRusa
                    Call Events_AI_DagaRusa(NpcIndex)
            End Select

        End With

        Exit Sub

ErrorHandler:
    
136     Call LogError("NPC.AI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)

138     Dim MiNPC As Npc: MiNPC = Npclist(NpcIndex)
    
140     Call QuitarNPC(NpcIndex)
142     Call RespawnNpc(MiNPC)

End Sub

Private Sub PerseguirUsuarioCercano(ByVal NpcIndex As Integer)

        On Error GoTo ErrorHandler

        Dim i                         As Long

        Dim UserIndex                 As Integer

        Dim UserIndexFront            As Integer

        Dim npcEraPasivo              As Boolean

        Dim agresor                   As Integer

        Dim minDistancia              As Integer

        Dim minDistanciaAtacable      As Integer

        Dim enemigoCercano            As Integer

        Dim enemigoAtacableMasCercano As Integer

        Dim distanciaOrigen           As Long
        
        ' Numero muy grande para que siempre haya un mÃƒÂ­nimo
100     minDistancia = 32000
102     minDistanciaAtacable = 32000

104     With Npclist(NpcIndex)
106         npcEraPasivo = .flags.OldHostil = 0
108         .Target = 0
110         .TargetNPC = 0

112         If .flags.AttackedBy <> vbNullString Then
114             agresor = NameIndex(.flags.AttackedBy)

            End If
            
            distanciaOrigen = Distancia(.Pos, .Orig)
            
            If UserIndex > 0 And UserIndexFront > 0 Then
            
                If NPCHasAUserInFront(NpcIndex, UserIndexFront) And EsEnemigo(NpcIndex, UserIndexFront) Then
                    enemigoAtacableMasCercano = UserIndexFront
                    minDistanciaAtacable = 1
                    minDistancia = 1

                End If

            Else

                ' Busco algun objetivo en el area.
                Dim query()    As Collision.UUID

                Dim TotalUsers As Integer

               ' Call ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
                
                For i = 0 To ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

                    UserIndex = query(i).Name
                    
                    If UserList(UserIndex).ConnIDValida Then

120                     If EsObjetivoValido(NpcIndex, UserIndex) Then

                            ' Busco el mas cercano, sea atacable o no.
122                         If Distancia(UserList(UserIndex).Pos, .Pos) < minDistancia And Not (UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Oculto) Then
124                             enemigoCercano = UserIndex
126                             minDistancia = Distancia(UserList(UserIndex).Pos, .Pos)

                            End If
                            
                            ' Busco el mas cercano que sea atacable.
128                         If (UsuarioAtacableConMagia(UserIndex) Or UsuarioAtacableConMelee(NpcIndex, UserIndex)) And Distancia(UserList(UserIndex).Pos, .Pos) < minDistanciaAtacable Then
130                             enemigoAtacableMasCercano = UserIndex
132                             minDistanciaAtacable = Distancia(UserList(UserIndex).Pos, .Pos)

                            End If
        
                        End If

                    End If
        
134             Next i

            End If

            ' Al terminar el `for`, puedo tener un maximo de tres objetivos distintos.
            ' Por prioridad, vamos a decidir estas cosas en orden.
            If distanciaOrigen < 40 Then
136             If npcEraPasivo Then

                    ' Significa que alguien le pego, y esta en modo agresivo trantando de darle.
                    ' El unico objetivo que importa aca es el atacante; los demas son ignorados.
138                 If EnRangoVision(NpcIndex, agresor) Then .Target = agresor
    
                Else ' El NPC es hostil siempre, le quiere pegar a alguien.
    
140                 If minDistanciaAtacable > 0 And enemigoAtacableMasCercano > 0 Then ' Hay alguien atacable cerca
142                     .Target = enemigoAtacableMasCercano
144                 ElseIf enemigoCercano > 0 Then ' Hay alguien cerca, pero no es atacable
146                     .Target = enemigoCercano

                    End If
    
                End If

            End If

            ' Si el NPC tiene un objetivo
148         If .Target > 0 And EsObjetivoValido(NpcIndex, .Target) Then
                    
                'asignamos heading nuevo al NPC seg�n el Target del nuevo usuario: .Char.Heading, si la distancia es <= 1
                If (.flags.Inmovilizado + .flags.Paralizado = 0) Then
                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, GetHeadingFromWorldPos(.Pos, UserList(.Target).Pos))

                End If

150             Call AI_AtacarUsuarioObjetivo(NpcIndex)
                    
                'Si se aleja mucho saca el target y empieza a volver a casa
                If distanciaOrigen > 60 Then .Target = 0
            Else

152             If .NPCtype <> eNPCType.GuardiaReal And .NPCtype <> eNPCType.GuardiasCaos Then


154                 Call RestoreOldMovement(NpcIndex)
                    ' No encontro a nadie cerca, camina unos pasos en cualquier direccion.
156                 Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)

                    ' # Se fija si se puede curar ?�
                    Call NpcLanzaUnSpell(NpcIndex)
                   
                Else

158                 If distanciaOrigen > 0 Then
160                     Call AI_CaminarConRumbo(NpcIndex, .Orig)
                    Else

162                     If .Char.Heading <> eHeading.SOUTH Then
164                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, eHeading.SOUTH)

                        End If

                    End If

                End If

            End If

        End With

        Exit Sub

ErrorHandler:

End Sub

' Cuando un NPC no tiene target y se puede mover libremente pero cerca de su lugar de origen.
' La mayoria de los NPC deberian mantenerse cerca de su posicion de origen, algunos quedaran quietos
' en su posicion y otros se moveran libremente cerca de su posicion de origen.
Private Sub AI_CaminarSinRumboCercaDeOrigen(ByVal NpcIndex As Integer)
        On Error GoTo AI_CaminarSinRumboCercaDeOrigen_Err

100     With Npclist(NpcIndex)
102         If .flags.Paralizado > 0 Or .flags.Inmovilizado > 0 Then
104             Call AnimacionIdle(NpcIndex, True)
106         ElseIf Distancia(.Pos, .Orig) > 4 Then
108             Call AI_CaminarConRumbo(NpcIndex, .Orig)
110         ElseIf RandomNumber(1, 6) = 3 Then
112             Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            Else
114             Call AnimacionIdle(NpcIndex, True)
            End If

        End With

        Exit Sub

AI_CaminarSinRumboCercaDeOrigen_Err:
        
End Sub

' Cuando un NPC no tiene target y se tiene que mover libremente
Private Sub AI_CaminarSinRumbo(ByVal NpcIndex As Integer)

        On Error GoTo AI_CaminarSinRumbo_Err

100     With Npclist(NpcIndex)

102         If RandomNumber(1, 6) = 3 And .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then
104             Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            Else
106             Call AnimacionIdle(NpcIndex, True)

            End If

        End With

        Exit Sub

AI_CaminarSinRumbo_Err:

End Sub

Private Sub AI_CaminarConRumbo(ByVal NpcIndex As Integer, ByRef rumbo As WorldPos)
        On Error GoTo AI_CaminarConRumbo_Err
    
100     If Npclist(NpcIndex).flags.Paralizado Or Npclist(NpcIndex).flags.Inmovilizado Then
102         Call AnimacionIdle(NpcIndex, True)
            Exit Sub
        End If
    
104     With Npclist(NpcIndex).pathFindingInfo
            ' Si no tiene un camino calculado o si el destino cambio
106         If .PathLength = 0 Or .Destination.X <> rumbo.X Or .Destination.Y <> rumbo.Y Then
108             .Destination.X = rumbo.X
110             .Destination.Y = rumbo.Y

                ' Recalculamos el camino
112             If SeekPath(NpcIndex, True) Then
                    ' Si consiguo un camino
114                 Call FollowPath(NpcIndex)
                End If
            Else ' Avanzamos en el camino
116             Call FollowPath(NpcIndex)
            End If

        End With

        Exit Sub

AI_CaminarConRumbo_Err:
        Dim errorDescription As String
118     errorDescription = Err.description & vbNewLine & " NpcIndex: " & NpcIndex & " NPCList.size= " & UBound(Npclist)

End Sub
Private Function NpcLanzaSpellInmovilizado(ByVal NpcIndex As Integer, ByVal tIndex As Integer) As Boolean
        
    NpcLanzaSpellInmovilizado = False
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado + .flags.Paralizado > 0 Then
            Select Case .Char.Heading
                Case eHeading.NORTH
                    If .Pos.X = UserList(tIndex).Pos.X And .Pos.Y > UserList(tIndex).Pos.Y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                    
                Case eHeading.EAST
                    If .Pos.Y = UserList(tIndex).Pos.Y And .Pos.X < UserList(tIndex).Pos.X Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                
                Case eHeading.SOUTH
                    If .Pos.X = UserList(tIndex).Pos.X And .Pos.Y < UserList(tIndex).Pos.Y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                
                Case eHeading.WEST
                    If .Pos.Y = UserList(tIndex).Pos.Y And .Pos.X > UserList(tIndex).Pos.X Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
            End Select
        Else
            NpcLanzaSpellInmovilizado = True
        End If
    End With
    
End Function

Public Function ComputeNextHeadingPos(ByVal NpcIndex As Integer) As WorldPos
On Error Resume Next
With Npclist(NpcIndex)
    ComputeNextHeadingPos.Map = .Pos.Map
    ComputeNextHeadingPos.X = .Pos.X
    ComputeNextHeadingPos.Y = .Pos.Y
    
    Select Case .Char.Heading
        Case eHeading.NORTH
            ComputeNextHeadingPos.Y = ComputeNextHeadingPos.Y - 1
        Exit Function
        
        Case eHeading.SOUTH
            ComputeNextHeadingPos.Y = ComputeNextHeadingPos.Y + 1
        Exit Function
        
        Case eHeading.EAST
            ComputeNextHeadingPos.X = ComputeNextHeadingPos.X + 1
        Exit Function
        
        Case eHeading.WEST
            ComputeNextHeadingPos.X = ComputeNextHeadingPos.X - 1
        Exit Function
        
    End Select
End With
End Function

Public Function NPCHasAUserInFront(ByVal NpcIndex As Integer, ByRef UserIndex As Integer) As Boolean
    On Error Resume Next
    Dim NextPosNPC As WorldPos
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        NPCHasAUserInFront = False
        Exit Function
    End If
    
    
    
    NextPosNPC = ComputeNextHeadingPos(NpcIndex)
    UserIndex = MapData(NextPosNPC.Map, NextPosNPC.X, NextPosNPC.Y).UserIndex
    NPCHasAUserInFront = (UserIndex > 0)
End Function


Private Sub AI_AtacarUsuarioObjetivo(ByVal AtackerNpcIndex As Integer)
        On Error GoTo ErrorHandler

        Dim AtacaConMagia As Boolean
        Dim AtacaMelee As Boolean
        Dim EstaPegadoAlUsuario As Boolean
        Dim tHeading As Byte
        Dim NextPosNPC As WorldPos
        Dim AtacaAlDelFrente As Boolean
        
        AtacaAlDelFrente = False
100     With Npclist(AtackerNpcIndex)
102         If .Target = 0 Then Exit Sub
                
              
104         EstaPegadoAlUsuario = (Distancia(.Pos, UserList(.Target).Pos) <= 1)
106         AtacaConMagia = .flags.LanzaSpells And _
                            modNuevoTimer.Intervalo_CriatureAttack(AtackerNpcIndex, False) And _
                            (RandomNumber(1, 100) <= 50)
             
108         AtacaMelee = EstaPegadoAlUsuario And UsuarioAtacableConMelee(AtackerNpcIndex, .Target) And .flags.Paralizado = 0
            AtacaMelee = AtacaMelee And (.flags.LanzaSpells > 0 And (UserList(.Target).flags.Invisible > 0 Or UserList(.Target).flags.Oculto > 0))
            AtacaMelee = AtacaMelee Or .flags.LanzaSpells = 0
            
            ' Se da vuelta y enfrenta al Usuario
109         tHeading = GetHeadingFromWorldPos(.Pos, UserList(.Target).Pos)
            
110         If AtacaConMagia Then
                ' Le lanzo un Hechizo
                If NpcLanzaSpellInmovilizado(AtackerNpcIndex, .Target) Then
                    Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)
112                 Call NpcLanzaUnSpell(AtackerNpcIndex)
                End If
114         ElseIf AtacaMelee Then
                Dim ChangeHeading As Boolean
                ChangeHeading = (.flags.Inmovilizado > 0 And tHeading = .Char.Heading) Or (.flags.Inmovilizado + .flags.Paralizado = 0)
                
                Dim UserIndexFront As Integer
                NextPosNPC = ComputeNextHeadingPos(AtackerNpcIndex)
                UserIndexFront = MapData(NextPosNPC.Map, NextPosNPC.X, NextPosNPC.Y).UserIndex
                AtacaAlDelFrente = (UserIndexFront > 0)
                
                If ChangeHeading Then
                    Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)
                End If
                
                If AtacaAlDelFrente And Not .flags.Paralizado = 1 Then
                    Call AnimacionIdle(AtackerNpcIndex, True)
                    If UserIndexFront > 0 Then
                        If UserList(UserIndexFront).flags.Muerto = 0 Then
                            If UserList(UserIndexFront).Faction.Status = 1 And (.NPCtype = eNPCType.GuardiaReal) Then
                                
                            Else
                                Call NpcAtacaUser(AtackerNpcIndex, UserIndexFront, tHeading)
                            End If
                        End If
                    End If

                End If
            End If

124         If UsuarioAtacableConMagia(.Target) Or UsuarioAtacableConMelee(AtackerNpcIndex, .Target) Then
                ' Si no tiene un camino pero esta pegado al usuario, no queremos gastar tiempo calculando caminos.
126             If .pathFindingInfo.PathLength = 0 And EstaPegadoAlUsuario Then Exit Sub
            
128             Call AI_CaminarConRumbo(AtackerNpcIndex, UserList(.Target).Pos)
            Else
130             Call AI_CaminarSinRumboCercaDeOrigen(AtackerNpcIndex)
            End If
        End With

        Exit Sub

ErrorHandler:

End Sub

Public Sub AI_GuardiaPersigueNpc(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        Dim targetPos As WorldPos
        
100     With Npclist(NpcIndex)
        
102          If .TargetNPC > 0 Then
104             targetPos = Npclist(.TargetNPC).Pos
                
106             If Distancia(.Pos, targetPos) <= 1 Then
108                 Call NpcAtacaNpc(NpcIndex, .TargetNPC, False)
                End If
                
110             If DistanciaRadial(.Orig, targetPos) <= (DIAMETRO_VISION_GUARDIAS_NPCS \ 2) Then
112                 If Npclist(.TargetNPC).Target = 0 Then
114                     Call AI_CaminarConRumbo(NpcIndex, targetPos)
116                 ElseIf UserList(Npclist(.TargetNPC).Target).flags.NPCAtacado <> .TargetNPC Then
118                     Call AI_CaminarConRumbo(NpcIndex, targetPos)
                    Else
120                     .TargetNPC = 0
122                     Call AI_CaminarConRumbo(NpcIndex, .Orig)
                    End If
                Else
124                 .TargetNPC = 0
126                 Call AI_CaminarConRumbo(NpcIndex, .Orig)
                End If
                
            Else
128             .TargetNPC = BuscarNpcEnArea(NpcIndex)
130             If Distancia(.Pos, .Orig) > 0 Then
132                 Call AI_CaminarConRumbo(NpcIndex, .Orig)
                Else
134                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, eHeading.SOUTH)
                End If
            End If
            
            
        End With
        
        Exit Sub
        
        
ErrorHandler:

End Sub

Private Function DistanciaRadial(OrigenPos As WorldPos, DestinoPos As WorldPos) As Long
100     DistanciaRadial = max(Abs(OrigenPos.X - DestinoPos.X), Abs(OrigenPos.Y - DestinoPos.Y))
End Function

Private Function BuscarNpcEnArea(ByVal NpcIndex As Integer) As Integer
        
        On Error GoTo BuscarNpcEnArea
        
        Dim X As Integer, Y As Integer
       
100    With Npclist(NpcIndex)
       
102         For X = (.Orig.X - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.X + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
104             For Y = (.Orig.Y - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.Y + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
                
106                 If MapData(.Orig.Map, X, Y).NpcIndex > 0 And NpcIndex <> MapData(.Orig.Map, X, Y).NpcIndex Then
                        Dim foundNpc As Integer
108                     foundNpc = MapData(.Orig.Map, X, Y).NpcIndex
                        
110                     If Npclist(foundNpc).Hostile Then
                        
112                         If Npclist(foundNpc).Target = 0 Then
114                             BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                                Exit Function
116                         ElseIf UserList(Npclist(foundNpc).Target).flags.NPCAtacado <> foundNpc Then
118                             BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                                Exit Function
                            End If
                            
                        End If
                        
                    End If
                    
120             Next Y
122         Next X

        End With
        
124     BuscarNpcEnArea = 0
        
        Exit Function

BuscarNpcEnArea:

End Function

Public Sub AI_NpcAtacaNpc(ByVal NpcIndex As Integer)

        On Error GoTo ErrorHandler
    
        Dim targetPos As WorldPos
        
        Dim Distance As Integer
        
        
100     With Npclist(NpcIndex)
            Distance = 3
            
102         If .TargetNPC > 0 Then
104             targetPos = Npclist(.TargetNPC).Pos
            
106             If InRangoVisionNPC(NpcIndex, targetPos.X, targetPos.Y) Then
                    
                    If .flags.Paralizado = 0 Then
                        ' Me fijo si el NPC esta al lado del Objetivo
                        
                        If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex)
                        Else
                            
108                         If Distancia(.Pos, targetPos) <= Distance Then
110                             Call NpcAtacaNpc(NpcIndex, .TargetNPC)

                            End If

                        End If
                  
                    End If
                  
112                 If .TargetNPC <> vbNull And .TargetNPC > 0 Then
114                     Call AI_CaminarConRumbo(NpcIndex, targetPos)

                    End If
               
                    Exit Sub

                End If

            End If
           
116         Call RestoreOldMovement(NpcIndex)
 
        End With
                
        Exit Sub
                
ErrorHandler:

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
        ' La IA que se ejecuta cuando alguien le pega al maestro de una Mascota/Elemental
        ' o si atacas a los NPCs con Movement = e_TipoAI.NpcDefensa
        ' A diferencia de IrUsuarioCercano(), aca no buscamos objetivos cercanos en el area
        ' porque ya establecemos como objetivo a el usuario que ataco a los NPC con este tipo de IA

        On Error GoTo SeguirAgresor_Err


100     If EsObjetivoValido(NpcIndex, Npclist(NpcIndex).Target) Then
102         Call AI_AtacarUsuarioObjetivo(NpcIndex)
        Else
104         Call RestoreOldMovement(NpcIndex)

        End If

        Exit Sub

SeguirAgresor_Err:

End Sub

Public Sub SeguirAmo(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        
100     With Npclist(NpcIndex)
        
102         If .MaestroUser = 0 Or Not .flags.Follow Then Exit Sub
        
            ' Si la mascota no tiene objetivo establecido.
104         If .Target = 0 And .TargetNPC = 0 Then
            
106             If EnRangoVision(NpcIndex, .MaestroUser) Then
108                 If UserList(.MaestroUser).flags.Muerto = 0 And _
                        UserList(.MaestroUser).flags.Invisible = 0 And _
                        UserList(.MaestroUser).flags.Oculto = 0 And _
                        Distancia(.Pos, UserList(.MaestroUser).Pos) > 3 Then
                    
                        ' Caminamos cerca del usuario
110                     Call AI_CaminarConRumbo(NpcIndex, UserList(.MaestroUser).Pos)
                        Exit Sub
                    
                    End If
                End If
                
112             Call AI_CaminarSinRumbo(NpcIndex)
            End If
        End With
    
        Exit Sub

ErrorHandler:
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

        On Error GoTo RestoreOldMovement_Err

100     With Npclist(NpcIndex)
102         .Target = 0
104         .TargetNPC = 0
        
            ' Si el NPC no tiene maestro, reseteamos el movimiento que tenia antes.
106         If .MaestroUser = 0 Then
108             .Movement = .flags.OldMovement
110             .Hostile = .flags.OldHostil
112             .flags.AttackedBy = vbNullString
            Else
            
                ' Si tiene maestro, hacemos que lo siga.
114             Call FollowAmo(NpcIndex)
            
            End If

        End With

        Exit Sub

RestoreOldMovement_Err:

End Sub

Private Sub HacerCaminata(ByVal NpcIndex As Integer)
        On Error GoTo Handler
    
        Dim Destino As WorldPos
        Dim Heading As eHeading
        Dim NextTile As WorldPos
        Dim MoveChar As Integer
        Dim PudoMover As Boolean

100     With Npclist(NpcIndex)
    
102         Destino.Map = .Pos.Map
104         Destino.X = .Orig.X + .Caminata(.CaminataActual).offset.X
106         Destino.Y = .Orig.Y + .Caminata(.CaminataActual).offset.Y

            ' Si todaviï�½a no llego al destino
108         If .Pos.X <> Destino.X Or .Pos.Y <> Destino.Y Then
        
                ' Tratamos de acercarnos (podemos pisar npcs, usuarios o triggers)
110             Heading = GetHeadingFromWorldPos(.Pos, Destino)

                ' Obtengo la posicion segun el heading
112             NextTile = .Pos
114             Call HeadtoPos(Heading, NextTile)
            
                ' Si hay un NPC
116             MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).NpcIndex
118             If MoveChar Then
                    ' Lo movemos hacia un lado
120                 Call MoveNpcToSide(MoveChar, Heading)
                End If
            
                ' Si hay un user
122             MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).UserIndex
124             If MoveChar Then
                    ' Si no esta muerto o es admin invisible (porque a esos los atraviesa)
126                 If UserList(MoveChar).flags.AdminInvisible = 0 Or UserList(MoveChar).flags.Muerto = 0 Then
                        ' Lo movemos hacia un lado
128                     Call MoveUserToSide(MoveChar, Heading)
                    End If
                End If
            
                ' Movemos al NPC de la caminata
130             PudoMover = MoveNPCChar(NpcIndex, Heading)
            
                ' Si no pudimos moverlo, hacemos como si hubiese llegado a destino... para evitar que se quede atascado
132             If Not PudoMover Or Distancia(.Pos, Destino) = 0 Then
            
                    ' Llegamos a destino, ahora esperamos el tiempo necesario para continuar
134                 .Contadores.Velocity = GetTime + .Caminata(.CaminataActual).Espera - .Velocity
                
                    ' Pasamos a la siguiente caminata
136                 .CaminataActual = .CaminataActual + 1
                
                    ' Si pasamos el ultimo, volvemos al primero
138                 If .CaminataActual > UBound(.Caminata) Then
140                     .CaminataActual = 1
                    End If
                
                End If
            
            ' Si por alguna razÃƒÂ³n estamos en el destino, seguimos con la siguiente caminata
            Else
        
142             .CaminataActual = .CaminataActual + 1
            
                ' Si pasamos el ultimo, volvemos al primero
144             If .CaminataActual > UBound(.Caminata) Then
146                 .CaminataActual = 1
                End If
            
            End If
    
        End With
    
        Exit Sub
    
Handler:

End Sub

Private Sub MovimientoInvasion(ByVal NpcIndex As Integer)
        On Error GoTo Handler
    
100     With Npclist(NpcIndex)
            Dim SpawnBox As t_SpawnBox
102         'SpawnBox = Invasiones(.flags.InvasionIndex).SpawnBoxes(.flags.SpawnBox)
    
            ' Calculamos la distancia a la muralla y generamos una posicion de destino
            Dim DistanciaMuralla As Integer, Destino As WorldPos
104         Destino = .Pos
        
106         If SpawnBox.Heading = eHeading.EAST Or SpawnBox.Heading = eHeading.WEST Then
108             DistanciaMuralla = Abs(.Pos.X - SpawnBox.CoordMuralla)
110             Destino.X = SpawnBox.CoordMuralla
            Else
112             DistanciaMuralla = Abs(.Pos.Y - SpawnBox.CoordMuralla)
114             Destino.Y = SpawnBox.CoordMuralla
            End If

            ' Si todavia esta lejos de la muralla
116         If DistanciaMuralla > 1 Then
        
                ' Tratamos de acercarnos (sin pisar)
                Dim Heading As eHeading
118             Heading = GetHeadingFromWorldPos(.Pos, Destino)
            
                ' Nos aseguramos que la posicion nueva esta dentro del rectangulo valido
                Dim NextTile As WorldPos
120             NextTile = .Pos
122             Call HeadtoPos(Heading, NextTile)
            
                ' Si la posicion nueva queda fuera del rectangulo valido
124             If Not InsideRectangle(SpawnBox.LegalBox, NextTile.X, NextTile.Y) Then
                    ' Invertimos la direccion de movimiento
126                 Heading = InvertHeading(Heading)
                End If
            
                ' Movemos el NPC
128             Call MoveNPCChar(NpcIndex, Heading)
        
            ' Si esta pegado a la muralla
            Else
        
                ' Chequeamos el intervalo de ataque
130             If Not Intervalo_CriatureAttack(NpcIndex, False) Then
                    Exit Sub
                End If
            
                ' Nos aseguramos que mire hacia la muralla
132             If .Char.Heading <> SpawnBox.Heading Then
134                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, SpawnBox.Heading)
                End If
            
                ' Sonido de ataque (si tiene)
136             If .flags.Snd1 > 0 Then
138                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(.flags.Snd1, .Pos.X, .Pos.Y))
                End If
            
                ' Sonido de impacto
140             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(SND_IMPACTO, .Pos.X, .Pos.Y))
            
                ' Da�amos la muralla
142             'Call HacerDa�oMuralla(.flags.InvasionIndex, RandomNumber(.Stats.MinHit, .Stats.MaxHit))  ' TODO: Defensa de la muralla? No hace falta creo...

            End If
    
        End With

        Exit Sub
    
Handler:
        Dim errorDescription As String
144     'errorDescription = Err.description & vbNewLine & "NpcId=" & Npclist(NpcIndex).Numero & " InvasionIndex:" & Npclist(NpcIndex).flags.InvasionIndex & " SpawnBox:" & Npclist(NpcIndex).flags.SpawnBox & vbNewLine
146     'Call TraceError(Err.Number, errorDescription, "AI.MovimientoInvasion", Erl)
End Sub

' El NPC elige un hechizo al azar dentro de su listado, con un potencial Target.
' Depdendiendo el tipo de spell que elije, se elije un target distinto que puede ser:
' - El .Target, el NPC mismo o area.
Private Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer)

        On Error GoTo NpcLanzaUnSpell_Err
        
        If Npclist(NpcIndex).flags.LanzaSpells = 0 Then Exit Sub
        ' Elegir hechizo, dependiendo del hechizo lo tiro sobre NPC, sobre Target o Sobre area (cerca de user o NPC si no tiene)
        Dim SpellIndex As Integer
        Dim Target     As Integer
        Dim PuedeDanarAlUsuario As Boolean

100     If Not Intervalo_CriatureAttack(NpcIndex, False) Then Exit Sub

102     Target = Npclist(NpcIndex).Target
104     SpellIndex = Npclist(NpcIndex).Spells(RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells))
106     PuedeDanarAlUsuario = Npclist(NpcIndex).flags.Paralizado = 0
        
        If SpellIndex = 0 Then Exit Sub
        
        
        
108     Select Case Hechizos(SpellIndex).Target

            Case TargetType.uUsuarios

110             If UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
112                 Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)

114                 If UserList(Target).flags.AtacadoPorNpc = 0 Then
116                     UserList(Target).flags.AtacadoPorNpc = NpcIndex

                    End If

                End If

118         Case TargetType.uNPC
120             If Hechizos(SpellIndex).AutoLanzar = 1 Then
122                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)

124             ElseIf Npclist(NpcIndex).TargetNPC > 0 Then
126                 Call NpcLanzaSpellSobreNpc(NpcIndex, Npclist(NpcIndex).TargetNPC, SpellIndex)

                End If

128         Case TargetType.uUsuariosYnpc

130             If Hechizos(SpellIndex).AutoLanzar = 1 Then
132                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)

134             ElseIf UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
136                 Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)

138                 If UserList(Target).flags.AtacadoPorNpc = 0 Then
140                     UserList(Target).flags.AtacadoPorNpc = NpcIndex

                    End If

142             ElseIf Npclist(NpcIndex).TargetNPC > 0 Then
144                 Call NpcLanzaSpellSobreNpc(NpcIndex, Npclist(NpcIndex).TargetNPC, SpellIndex)

                End If

146         Case TargetType.uTerreno
148             'Call NpcLanzaSpellSobreArea(NpcIndex, SpellIndex)

        End Select

        Exit Sub

NpcLanzaUnSpell_Err:


End Sub

Private Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
        On Error GoTo NpcLanzaUnSpellSobreNpc_Err
    
100     With Npclist(NpcIndex)
        
102         If Not Intervalo_CriatureAttack(NpcIndex, False) Then Exit Sub
104         If .Pos.Map <> Npclist(TargetNPC).Pos.Map Then Exit Sub
    
            Dim K As Integer
106             K = RandomNumber(1, .flags.LanzaSpells)

108         Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, .Spells(K))
    
        End With
     
        Exit Sub

NpcLanzaUnSpellSobreNpc_Err:

End Sub


' ---------------------------------------------------------------------------------------------------
'                                       HELPERS
' ---------------------------------------------------------------------------------------------------

Public Function EsObjetivoValido(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
100     If UserIndex = 0 Then Exit Function
        If NpcIndex = 0 Then Exit Function

        ' Esta condicion debe ejecutarse independiemente de el modo de busqueda.
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Function 'User muerto
102     If Not EnRangoVision(NpcIndex, UserIndex) Then Exit Function 'En rango
        If Not EsEnemigo(NpcIndex, UserIndex) Then Exit Function 'Es enemigo
        If UserList(UserIndex).flags.EnConsulta = 1 Then Exit Function 'En consulta
        If EsGm(UserIndex) And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
       
        
        EsObjetivoValido = True

End Function

Private Function EsEnemigo(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

        On Error GoTo EsEnemigo_Err


100     If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

102     With Npclist(NpcIndex)

104         If .flags.AttackedBy <> vbNullString Then
106             EsEnemigo = (UserIndex = NameIndex(.flags.AttackedBy))
108             If EsEnemigo Then Exit Function
            End If

110         Select Case .flags.AIAlineacion
                Case e_Alineacion.Real
112                 EsEnemigo = Escriminal(UserIndex)

114             Case e_Alineacion.Caos
116                 EsEnemigo = Not Escriminal(UserIndex)

118             Case e_Alineacion.ninguna
120                 EsEnemigo = True
                    ' Ok. No hay nada especial para hacer, cualquiera puede ser enemigo!

            End Select

        End With

        Exit Function

EsEnemigo_Err:

End Function

Private Function EnRangoVision(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

        On Error GoTo EnRangoVision_Err

        Dim userPos As WorldPos
        Dim NpcPos As WorldPos
        Dim Limite_x As Integer, Limite_y As Integer

        ' Si alguno es cero, devolve false
100     If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

102     Limite_x = IIf(Npclist(NpcIndex).Distancia <> 0, Npclist(NpcIndex).Distancia, RANGO_VISION_x)
104     Limite_y = IIf(Npclist(NpcIndex).Distancia <> 0, Npclist(NpcIndex).Distancia, RANGO_VISION_y)

106     userPos = UserList(UserIndex).Pos
108     NpcPos = Npclist(NpcIndex).Pos

110     EnRangoVision = ( _
          (userPos.Map = NpcPos.Map) And _
          (Abs(userPos.X - NpcPos.X) <= Limite_x) And _
          (Abs(userPos.Y - NpcPos.Y) <= Limite_y) _
        )


        Exit Function

EnRangoVision_Err:
End Function

Private Function UsuarioAtacableConMagia(ByVal targetUserIndex As Integer) As Boolean

        On Error GoTo UsuarioAtacableConMagia_Err

100     If targetUserIndex = 0 Then Exit Function

102     With UserList(targetUserIndex)
104       UsuarioAtacableConMagia = ( _
            .flags.Muerto = 0 And _
            .flags.Invisible = 0 And _
            .flags.Oculto = 0 And _
            .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And _
            Not (EsGm(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And _
            Not .flags.EnConsulta)
        End With


        Exit Function

UsuarioAtacableConMagia_Err:

End Function

Private Function UsuarioAtacableConMelee(ByVal NpcIndex As Integer, ByVal targetUserIndex As Integer) As Boolean

        On Error GoTo UsuarioAtacableConMelee_Err

100     If targetUserIndex = 0 Then Exit Function

        Dim EstaPegadoAlUser As Boolean
    
102     With UserList(targetUserIndex)
    
104       EstaPegadoAlUser = Distancia(Npclist(NpcIndex).Pos, .Pos) = 1

106       UsuarioAtacableConMelee = ( _
            .flags.Muerto = 0 And _
            (EstaPegadoAlUser Or (Not EstaPegadoAlUser And (.flags.Invisible + .flags.Oculto) = 0)) And _
            .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And _
            Not (EsGm(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And _
            Not .flags.EnConsulta)
        End With

        Exit Function

UsuarioAtacableConMelee_Err:

End Function



