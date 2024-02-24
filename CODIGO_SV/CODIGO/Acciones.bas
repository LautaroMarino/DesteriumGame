Attribute VB_Name = "Acciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, _
           ByVal Map As Integer, _
           ByVal X As Integer, _
           ByVal Y As Integer, _
           ByVal Tipo As Byte)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo Accion_Err
        '</EhHeader>

        Dim TempIndex As Integer
    
        '¿Rango Visión? (ToxicWaste)
100     If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_x) Then

            Exit Sub

        End If
    
        '¿Posicion valida?
102     If InMapBounds(Map, X, Y) Then

104         With UserList(UserIndex)

106             If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
108                 TempIndex = MapData(Map, X, Y).NpcIndex
                
                    'Set the target NPC
110                 .flags.TargetNPC = TempIndex
                
112                 If (Npclist(TempIndex).Comercia = 1 And (Tipo = 1 Or Tipo = 0)) Then

                        '¿Esta el user muerto? Si es asi no puede comerciar
114                     If .flags.Muerto = 1 Then
116                         Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If

                    
                        'Is it already in commerce mode??
118                     If .flags.Comerciando Then

                            Exit Sub

                        End If
                    
120                     If Distancia(Npclist(TempIndex).Pos, .Pos) > 5 Then
122                         Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                    
                        'Iniciamos la rutina pa' comerciar.
124                     Call IniciarComercioNPC(UserIndex)
                    
126                 ElseIf (Npclist(TempIndex).Quest > 0 And (Tipo = 2 Or Tipo = 0)) Then
                        '¿Esta el user muerto? Si es asi no puede comerciar
128                     If .flags.Muerto = 1 Then
130                         Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
    
                            Exit Sub
    
                        End If
                    
                        'Is it already in commerce mode??
132                     If .flags.Comerciando Then

                            Exit Sub

                        End If
                    
134                     If Distancia(Npclist(TempIndex).Pos, .Pos) > 5 Then
136                         Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                    
138                     Call WriteViewListQuest(UserIndex, Npclist(TempIndex).Quests, Npclist(TempIndex).Name)
                    
168                 ElseIf Npclist(TempIndex).NPCtype = eNPCType.Revividor Or Npclist(TempIndex).NPCtype = eNPCType.ResucitadorNewbie Then

170                     If Distancia(.Pos, Npclist(TempIndex).Pos) > 10 Then
172                         Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                    
                        'Revivimos si es necesario
174                     If .flags.Muerto = 1 And (Npclist(TempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
176                         Call RevivirUsuario(UserIndex)
                        End If
                    
178                     If Npclist(TempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                            'curamos totalmente
180                         .Stats.MinHp = .Stats.MaxHp
182                         Call WriteUpdateUserStats(UserIndex)
                        End If
                    
184                 ElseIf Npclist(TempIndex).NPCtype = eNPCType.Fundition Then
                    
                    ElseIf Npclist(TempIndex).NPCtype = eNPCType.Mascota Then
                        If .MascotaIndex = TempIndex Then
                            Call QuitarPet(UserIndex, .MascotaIndex)
                            Exit Sub
                
                        End If
192                 ElseIf Npclist(TempIndex).numero = TRAVEL_NPC_HOME Then

                       ' If Distancia(.Pos, Npclist(TempIndex).Pos) > 2 Then
                         '   Call WriteConsoleMsg(UserIndex, "Acercate más y te llevaré de regreso.", FontTypeNames.FONTTYPE_INFO)

                           ' Exit Sub

                       ' End If
                    
                        'Dim Pos As WorldPos
                    
                       ' Pos.Map = Ullathorpe.Map
                       ' Pos.X = Ullathorpe.X
                      '  Pos.Y = Ullathorpe.Y
                    
                        'ClosestStablePos Pos, Pos
                       ' Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, True)
                    End If
                
                    '¿Es un obj?
194             ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
196                 TempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
198                 .flags.TargetObj = TempIndex
                
200                 Select Case ObjData(TempIndex).OBJType

                        Case eOBJType.otPuertas 'Es una puerta
202                         Call AccionParaPuerta(Map, X, Y, UserIndex)
                    
204                     Case eOBJType.otcofre ' Cofres cerrados tirados por el mundo
206                         Call AccionParaCofre(Map, X, Y, UserIndex)
                    End Select

                    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
208             ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
210                 TempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
212                 .flags.TargetObj = TempIndex
                
214                 Select Case ObjData(TempIndex).OBJType
                    
                        Case eOBJType.otPuertas 'Es una puerta
216                         Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                    
                    End Select
            
218             ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
220                 TempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
222                 .flags.TargetObj = TempIndex
        
224                 Select Case ObjData(TempIndex).OBJType

                        Case eOBJType.otPuertas 'Es una puerta
226                         Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
                    End Select
            
228             ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
230                 TempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
232                 .flags.TargetObj = TempIndex
                
234                 Select Case ObjData(TempIndex).OBJType

                        Case eOBJType.otPuertas 'Es una puerta
236                         Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
                    End Select

                End If

            End With

        End If

        '<EhFooter>
        Exit Sub

Accion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Acciones.Accion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo AccionParaPuerta_Err
        '</EhHeader>
    
        Dim ObjIndex As Integer
    
100     If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
102         If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
104             If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                    'Abre la puerta
106                 If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
108                     MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
110                     ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
112                     Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjIndex, ObjData(ObjIndex).GrhIndex, X, Y, ObjData(ObjIndex).Name, 0, ObjData(ObjIndex).Sound))
                    
                        'Desbloquea
114                     MapData(Map, X, Y).Blocked = 0
116                     MapData(Map, X - 1, Y).Blocked = 0
                    
                        'Bloquea todos los mapas
118                     Call Bloquear(True, Map, X, Y, 0)
120                     Call Bloquear(True, Map, X - 1, Y, 0)
                      
                        'Sonido
122                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PUERTA, X, Y))
                    
                    Else
124                     Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else
                    'Cierra puerta
126                 MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
128                 ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
130                 Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjIndex, ObjData(ObjIndex).GrhIndex, X, Y, ObjData(ObjIndex).Name, 0, ObjData(ObjIndex).Sound))
                                
132                 MapData(Map, X, Y).Blocked = 1
134                 MapData(Map, X - 1, Y).Blocked = 1
                
136                 Call Bloquear(True, Map, X - 1, Y, 1)
138                 Call Bloquear(True, Map, X, Y, 1)
                
140                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PUERTA, X, Y))
                End If
        
142             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
            Else
144             Call WriteConsoleMsg(UserIndex, "La puerta está cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
            End If

        Else
146         Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

AccionParaPuerta_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Acciones.AccionParaPuerta " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AccionParaCofre(ByVal Map As Integer, _
                            ByVal X As Integer, _
                            ByVal Y As Integer, _
                            ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo AccionParaCofre_Err
        '</EhHeader>
    
        Dim ObjIndex   As Integer

        Dim Obj        As ObjData

        Dim ObjAbierto As Obj
    
        Dim DropObj As Boolean
    
        Dim Time As Double
    
100     Time = GetTime
102     DropObj = True

          ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
106     Obj = ObjData(ObjIndex)
    
108     If UserList(UserIndex).flags.Muerto Then
110         Call WriteConsoleMsg(UserIndex, "¡No has logrado abrir el cofre!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
     If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2 Then
         Call WriteConsoleMsg(UserIndex, "¡No puedes abrir el cofre desde lejos!", FontTypeNames.FONTTYPE_INFORED)
        
            Exit Sub
        End If
        
        
        If UserList(UserIndex).Stats.Elv < ObjData(ObjIndex).LvlMin Then
            Call WriteConsoleMsg(UserIndex, "¡Tu nivel no te permite abrir el Cofre!", FontTypeNames.FONTTYPE_INFORED)
        
            Exit Sub
        End If
        
116     If (Time - MapData(Map, X, Y).TimeClic) < (Obj.Chest.ClicTime * 1000) Then
118         Call WriteConsoleMsg(UserIndex, "¡Haz forzado abrir el cofre antes de tiempo! Debes esperar un poco más...", FontTypeNames.FONTTYPE_INFORED)
        
            Exit Sub
        End If
    
        ' Probabilidad de que el cofre se abra y vuelva a cerrarse
120     If RandomNumber(1, 100) <= Obj.Chest.ProbClose Then
122         Call WriteConsoleMsg(UserIndex, "No has contado con la suficiente fuerza para abrir el cofre. ¡Se ha cerrado!", FontTypeNames.FONTTYPE_INFORED)
124         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sChestClose, X, Y))
126         MapData(Map, X, Y).TimeClic = GetTime
            Exit Sub
        End If
                
        ' Probabilidad de que el cofre se abra y se rompa
128     If RandomNumber(1, 100) <= Obj.Chest.ProbBreak Then
130         Call WriteConsoleMsg(UserIndex, "Parece ser que el cofre se ha roto. ¡Tardará en reconstruirse!", FontTypeNames.FONTTYPE_INFORED)
132         DropObj = False ' Se rompe el Cofre
134         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sChestBreak, X, Y))
            GoTo Chesting:
        End If
    
   


Chesting:
    
        ' Ponemos el Cofre Abierto/Roto
136     If ChestData_Add(Map, X, Y, ObjIndex, Obj.Chest.RespawnTime, DropObj) Then
138         If DropObj Then Call Chest_DropObj(UserIndex, ObjIndex, Map, X, Y, False)
        
            ' Chequeamos las Quests
            Call Quests_AddChest(UserIndex, ObjIndex, 1)
            
            ' Quitamos el Cofre Cerrado
140         Call EraseObj(MapData(Map, X, Y).ObjInfo.Amount, Map, X, Y)
        
144         If DropObj Then

146             ObjAbierto.Amount = 1
148             ObjAbierto.ObjIndex = ObjData(ObjIndex).IndexAbierta    ' Cofre Abierto
            Else
        
150             ObjAbierto.Amount = 1
152             ObjAbierto.ObjIndex = ObjData(ObjIndex).IndexCerrada ' Cofre Roto
            End If
        
154         Call MakeObj(ObjAbierto, Map, X, Y)
        
        End If
                
        '<EhFooter>
        Exit Sub

AccionParaCofre_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Acciones.AccionParaCofre " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
