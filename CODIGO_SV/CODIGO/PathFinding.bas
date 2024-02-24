Attribute VB_Name = "PathFinding"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

'#######################################################
'PathFinding Module
'Coded By Gulfas Morgolock
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'
'Ore is an excellent engine for introducing you not only
'to online game programming but also to general
'game programming. I am convinced that Aaron Perkings, creator
'of ORE, did a great work. He made possible that a lot of
'people enjoy for no fee games made with his engine, and
'for me, this is something great.
'
'I'd really like to contribute to this work, and all the
'projects of free ore-based MMORPGs that are on the net.
'
'I did some basic improvements on the AI of the NPCs, I
'added pathfinding, so now, the npcs are able to avoid
'obstacles. I believe that this improvement was essential
'for the engine.
'
'I'd like to see this as my contribution to ORE project,
'I hope that someone finds this source code useful.
'So, please feel free to do whatever you want with my
'pathfinging module.
'
'I'd really appreciate that if you find this source code
'useful you mention my nickname on the credits of your
'program. But there is no obligation ;).
'
'.........................................................
'Note:
'There is a little problem, ORE refers to map arrays in a
'different manner that my pathfinding routines. When I wrote
'these routines, I did it without thinking in ORE, so in my
'program I refer to maps in the usual way I do it.
'
'For example, suppose we have:
'Map(1 to Y,1 to X) as MapBlock
'I usually use the first coordinate as Y, and
'the second one as X.
'
'ORE refers to maps in converse way, for example:
'Map(1 to X,1 to Y) as MapBlock. As you can see the
'roles of first and second coordinates are different
'that my routines
'
'.........................................................

'###########################################################################
' CHANGES
'
' 27/03/2021 WyroX: Fixed inverted coordinates and changed algorithm to A*
'###########################################################################


Option Explicit

Private Type t_IntermidiateWork
    Closed As Boolean
    Distance As Integer
    Previous As Position
    EstimatedTotalDistance As Single
End Type

Private OpenVertices(1000) As Position
Private VertexCount As Integer

Private Table(1 To 1432, 1 To 1780) As t_IntermidiateWork

Private DirOffset(eHeading.NORTH To eHeading.WEST) As Position

Private ClosestVertex As Position
Private ClosestDistance As Single

Private Const MAXINT As Integer = 32767

' WyroX: Usada para mover memoria... VB6 es un desastre en cuanto a contenedores din�micos
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal Length As Long)

Public Sub InitPathFinding()
        
        On Error GoTo InitPathFinding_Err

        Dim Heading As eHeading, DirH As Integer
        
100     For Heading = eHeading.NORTH To eHeading.WEST
105         DirOffset(Heading).X = (2 - DirH) * (DirH Mod 2)
110         DirOffset(Heading).Y = (DirH - 1) * (1 - (DirH Mod 2))
115         DirH = DirH + 1
        Next

        Exit Sub

InitPathFinding_Err:

End Sub

Public Sub FollowPath(ByVal NpcIndex As Integer)
        
        On Error GoTo FollowPath_Err
        
        Dim nextPos As WorldPos
    
100     With Npclist(NpcIndex)
            
            If (.pathFindingInfo.PathLength > UBound(.pathFindingInfo.Path)) Then ' Fix temporal para que no explote el LOG
                .pathFindingInfo.PathLength = 0
                Exit Sub
            End If
            
105         nextPos.Map = .Pos.Map
110         nextPos.X = .pathFindingInfo.Path(.pathFindingInfo.PathLength).X
115         nextPos.Y = .pathFindingInfo.Path(.pathFindingInfo.PathLength).Y
        
120         Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.Pos, nextPos))
125         .pathFindingInfo.PathLength = .pathFindingInfo.PathLength - 1
    
        End With
      
        Exit Sub

FollowPath_Err:

End Sub

Private Function InsideLimits(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo InsideLimits_Err
        
100     InsideLimits = X >= 1 And X <= XMaxMapSize And Y >= 1 And Y <= YMaxMapSize
        
        Exit Function

InsideLimits_Err:

End Function

Private Function IsWalkable(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As eHeading) As Boolean
        
    On Error GoTo ErrHandler
    
    Dim Map As Integer
1    Map = Npclist(NpcIndex).Pos.Map
    
    With MapData(Map, X, Y)

        ' Otro NPC
2        If .NpcIndex Then Exit Function
        
        ' Usuario
3        If .UserIndex And .UserIndex <> Npclist(NpcIndex).Target Then Exit Function

        ' Traslado
4        If .TileExit.Map <> 0 Then Exit Function

        ' Agua
5        If HayAgua(Map, X, Y) Then
            If Npclist(NpcIndex).flags.AguaValida = 0 Then Exit Function
            
        ' Tierra
        Else
6            If Npclist(NpcIndex).flags.TierraInvalida <> 0 Then Exit Function
        End If
        
        ' Trigger inv�lido para NPCs
        If .trigger = eTrigger.POSINVALIDA Then
            ' Si no es mascota
8            If Npclist(NpcIndex).MaestroNpc = 0 Then Exit Function
        End If
    
        ' Tile bloqueado
        If Npclist(NpcIndex).NPCtype <> eNPCType.GuardiaReal And Npclist(NpcIndex).NPCtype <> eNPCType.GuardiasCaos Then
9            If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function
            End If
        Else
10            If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, X + 1, Y) And Not HayPuerta(Map, X, Y) And Not HayPuerta(Map, X + 1, Y - 1) And Not HayPuerta(Map, X, Y - 1) Then Exit Function
        End If
            
    End With
    
11  IsWalkable = True
    
    Exit Function
    
ErrHandler:
End Function

Private Sub ProcessAdjacent(ByVal NpcIndex As Integer, ByVal CurX As Integer, ByVal CurY As Integer, ByVal Heading As eHeading, ByRef EndPos As Position)

    On Error GoTo ErrHandler
    
    Dim X As Integer, Y As Integer, DistanceFromStart As Integer, EstimatedDistance As Single
    
    With DirOffset(Heading)
1        X = CurX + .X
2        Y = CurY + .Y
    End With
    
    With Table(X, Y)

        ' Si ya est� cerrado, salimos
        If .Closed Then Exit Sub
    
        ' Nos quedamos en el campo de visi�n del NPC
        If InsideLimits(Npclist(NpcIndex).Pos.Map, X, Y) Then
        
            ' Si puede atravesar el tile al siguiente
3            If IsWalkable(NpcIndex, X, Y, Heading) Then
            
                ' Calculamos la distancia hasta este v�rtice
4                DistanceFromStart = Table(CurX, CurY).Distance + 1
    
                ' Si no hab�amos visitado este v�rtice
                If .Distance = MAXINT Then
                    ' Lo metemos en la cola
5                    Call OpenVertex(X, Y)
                    
                ' Si ya lo hab�amos visitado, nos fijamos si este camino es m�s corto
                ElseIf DistanceFromStart > .Distance Then
                    ' Es m�s largo, salimos
                    Exit Sub
                End If
    
                ' Guardamos la distancia desde el inicio
6                .Distance = DistanceFromStart
                
                ' La distancia estimada al objetivo
7                EstimatedDistance = EuclideanDistance(X, Y, EndPos)
                
                ' La distancia total estimada
8                .EstimatedTotalDistance = DistanceFromStart + EstimatedDistance
                
                ' Y la posici�n de la que viene
9                .Previous.X = CurX
10                .Previous.Y = CurY
                
                ' Si la distancia total estimada es la menor hasta ahora
                If EstimatedDistance < ClosestDistance Then
11                    ClosestDistance = EstimatedDistance
12                    ClosestVertex.X = X
13                    ClosestVertex.Y = Y
                End If
                
            End If
            
        End If

    End With
    
    Exit Sub
    
ErrHandler:
End Sub

Public Function SeekPath(ByVal NpcIndex As Integer, Optional ByVal Closest As Boolean) As Boolean
        ' Busca un camino desde la posici�n del NPC a la posici�n en .pathFindingInfo.Target
        ' El par�metro Closest indica que en caso de que no exista un camino completo, se debe retornar el camino parcial hasta la posici�n m�s cercana al objetivo.
        ' Si Closest = True, la funci�n devuelve True si puede moverse al menos un tile. Si Closest = False, devuelve True si se encontr� un camino completo.
        ' El camino se almacena en .pathFindingInfo.Path
        
        On Error GoTo SeekPath_Err
        
        Dim PosNPC As Position
        Dim PosTarget As Position
        Dim Heading As eHeading, Vertex As Position
        Dim MaxDistance As Integer, Index As Integer
        Dim MinTotalDistance As Integer, BestVertexIndex As Integer
        Dim UserIndex As Integer 'no es necesario
        Dim pasos As Long
        
        pasos = 0
        'Ya estamos en la posici�n.
        If UserIndex > 0 Then
            If NPCHasAUserInFront(NpcIndex, UserIndex) Then
                SeekPath = False
                Exit Function
            End If
        End If
        
        
100     With Npclist(NpcIndex)
105         PosNPC.X = .Pos.X
110         PosNPC.Y = .Pos.Y
    
            ' Posici�n objetivo
115         PosTarget.X = .pathFindingInfo.Destination.X
120         PosTarget.Y = .pathFindingInfo.Destination.Y
            
            ' Inicializar contenedores para el algoritmo
125         Call InitializeTable(Table, .Pos.Map, PosNPC, .pathFindingInfo.RangoVision)
130         VertexCount = 0
        
            ' A�adimos la posici�n inicial a la lista
135         Call OpenVertexV(PosNPC)
        
            ' Distancia m�xima a calcular (distancia en tiles al target + inteligencia del NPC)
140         MaxDistance = TileDistance(PosNPC, PosTarget) + .pathFindingInfo.Inteligencia
        
            ' Distancia euclideana desde la posici�n inicial hasta la final
145         Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance = EuclideanDistanceV(PosNPC, PosTarget)
            
            ' Ya estamos en la posicion
            If (Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance = 0) Then
                SeekPath = False
                Exit Function
            End If
            
            ' Distancia posici�n inicial
150         Table(PosNPC.X, PosNPC.Y).Distance = 0
        
            ' Distancia m�nima
155         ClosestDistance = Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance
160         ClosestVertex.X = PosNPC.X
165         ClosestVertex.Y = PosNPC.Y
        
        End With

        ' Loop principal del algoritmo
170     Do While (VertexCount > 0 And pasos < 300)
            
            pasos = pasos + 1
175         MinTotalDistance = MAXINT
        
            ' Buscamos en la cola la posici�n con menor distancia total
180         For Index = 0 To VertexCount - 1
        
185             With OpenVertices(Index)
            
190                 If Table(.X, .Y).EstimatedTotalDistance < MinTotalDistance Then
195                     MinTotalDistance = Table(.X, .Y).EstimatedTotalDistance
200                     BestVertexIndex = Index
                    End If
                
                End With
            
            Next
        
205         Vertex = OpenVertices(BestVertexIndex)
210         With Vertex
                ' Si es la posici�n objetivo
215             If .X = PosTarget.X And .Y = PosTarget.Y Then
            
                    ' Reconstru�mos el trayecto
220                 Call MakePath(NpcIndex, .X, .Y)
                
                    ' Salimos
225                 SeekPath = True
                    Exit Function
                
                End If

                ' Eliminamos la posici�n de la cola
230             Call CloseVertex(BestVertexIndex)

                ' Cerramos la posici�n actual
235             Table(.X, .Y).Closed = True

                ' Si a�n podemos seguir procesando m�s lejos
240             If Table(.X, .Y).Distance < MaxDistance Then
            
                    ' Procesamos adyacentes
245                 For Heading = eHeading.NORTH To eHeading.WEST
250                     Call ProcessAdjacent(NpcIndex, .X, .Y, Heading, PosTarget)
                    Next
                
                End If
            
            End With
        
        Loop
    
        ' No hay m�s nodos por procesar. O bien no existe un camino v�lido o el NPC no es suficientemente inteligente.
    
        ' Si debemos retornar la posici�n m�s cercana al objetivo
255     If Closest Then
    
            ' Si se recorri� al menos un tile
260         If ClosestVertex.X <> PosNPC.X Or ClosestVertex.Y <> PosNPC.Y Then
        
                ' Reconstru�mos el camino desde la posici�n m�s cercana al objetivo
265             Call MakePath(NpcIndex, ClosestVertex.X, ClosestVertex.Y)
            
270             SeekPath = True
                Exit Function
            
            End If
        
        End If

        ' Llegados a este punto, invalidamos el Path del NPC
275     Npclist(NpcIndex).pathFindingInfo.PathLength = 0

        Exit Function

SeekPath_Err:
End Function

Private Sub MakePath(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo MakePath_Err
 
100     With Npclist(NpcIndex)
            ' Obtenemos la distancia total del camino
105         .pathFindingInfo.PathLength = Table(X, Y).Distance

            Dim step As Integer
        
            ' Asignamos las coordenadas del resto camino, el final queda al inicio del array
110         For step = 1 To UBound(.pathFindingInfo.Path) ' .pathFindingInfo.PathLength TODO
        
115             With .pathFindingInfo.Path(step)
120                 .X = X
125                 .Y = Y
                End With
                If X > 0 And Y > 0 Then
130                 With Table(X, Y)
135                     X = .Previous.X
140                     Y = .Previous.Y
                    End With
                End If
            
            Next

        End With
   
        
        Exit Sub

MakePath_Err:
End Sub

Private Sub InitializeTable(ByRef Table() As t_IntermidiateWork, ByVal Map As Integer, ByRef PosNPC As Position, ByVal RangoVision As Single)
        ' Inicializar la tabla de posiciones para calcular el camino.
        ' Solo limpiamos el campo de visi�n del NPC.
        
        On Error GoTo InitializeTable_Err

        Dim X As Integer, Y As Integer

100     For Y = PosNPC.Y - RangoVision To PosNPC.Y + RangoVision
105         For X = PosNPC.X - RangoVision To PosNPC.X + RangoVision
        
110             If InsideLimits(Map, X, Y) Then
115                 Table(X, Y).Closed = False
120                 Table(X, Y).Distance = MAXINT
                End If
            
            Next
        Next

        
        Exit Sub

InitializeTable_Err:

End Sub

Private Function TileDistance(ByRef Vertex1 As Position, ByRef Vertex2 As Position) As Integer
        
        On Error GoTo TileDistance_Err
        
100     TileDistance = Abs(Vertex1.X - Vertex2.X) + Abs(Vertex1.Y - Vertex2.Y)
        
        Exit Function

TileDistance_Err:

End Function

Private Function EuclideanDistance(ByVal X As Integer, ByVal Y As Integer, ByRef Vertex As Position) As Single
        
        On Error GoTo EuclideanDistance_Err
        
        Dim dX As Integer, dY As Integer
100     dX = Vertex.X - X
105     dY = Vertex.Y - Y
110     EuclideanDistance = Sqr(dX * dX + dY * dY)
        
        Exit Function

EuclideanDistance_Err:

End Function

Private Function EuclideanDistanceV(ByRef Vertex1 As Position, ByRef Vertex2 As Position) As Single
        
        On Error GoTo EuclideanDistanceV_Err
        
        Dim dX As Integer, dY As Integer
100     dX = Vertex1.X - Vertex2.X
105     dY = Vertex1.Y - Vertex2.Y
110     EuclideanDistanceV = Sqr(dX * dX + dY * dY)
        
        Exit Function

EuclideanDistanceV_Err:

End Function

Private Sub OpenVertex(ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo OpenVertex_Err
        
100     With OpenVertices(VertexCount)
105         .X = X: .Y = Y
        End With
110     VertexCount = VertexCount + 1
        
        Exit Sub

OpenVertex_Err:

End Sub

Private Sub OpenVertexV(ByRef Vertex As Position)
        
        On Error GoTo OpenVertexV_Err
        
100     OpenVertices(VertexCount) = Vertex
105     VertexCount = VertexCount + 1
        
        Exit Sub

OpenVertexV_Err:

End Sub

Private Sub CloseVertex(ByVal Index As Integer)
        
        On Error GoTo CloseVertex_Err
        
100     VertexCount = VertexCount - 1
105     Call MoveMemory(OpenVertices(Index), OpenVertices(Index + 1), Len(OpenVertices(0)) * (VertexCount - Index))
        
        Exit Sub

CloseVertex_Err:

End Sub

' Las posiciones se pasan ByRef pero NO SE MODIFICAN.
Public Function GetHeadingFromWorldPos(ByRef CurrentPos As WorldPos, ByRef nextPos As WorldPos) As eHeading
        
        On Error GoTo GetHeadingFromWorldPos_Err
        
        Dim dX As Integer, dY As Integer
    
100     dX = nextPos.X - CurrentPos.X
105     dY = nextPos.Y - CurrentPos.Y
    
110     If dX < 0 Then
115         GetHeadingFromWorldPos = eHeading.WEST
120     ElseIf dX > 0 Then
125         GetHeadingFromWorldPos = eHeading.EAST
130     ElseIf dY < 0 Then
135         GetHeadingFromWorldPos = eHeading.NORTH
        Else
140         GetHeadingFromWorldPos = eHeading.SOUTH
        End If

        Exit Function

GetHeadingFromWorldPos_Err:

End Function

