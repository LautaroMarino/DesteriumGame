Attribute VB_Name = "NPCs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuitarMascota_Err
        '</EhHeader>

        Dim i As Integer
    
100     With UserList(UserIndex)

102         If .MascotaIndex Then
104             .MascotaIndex = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

QuitarMascota_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.QuitarMascota " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuitarMascotaNpc_Err
        '</EhHeader>

100     Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
        '<EhFooter>
        Exit Sub

QuitarMascotaNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.QuitarMascotaNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo MuereNpc_Err

        '</EhHeader>

        '********************************************************
        'Author: Unknown
        'Llamado cuando la vida de un NPC llega a cero.
        'Last Modify Date: 13/07/2010
        '22/06/06: (Nacho) Chequeamos si es pretoriano
        '24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
        '22/05/2010: ZaMa - Los caos ya no suben nobleza ni plebe al atacar npcs.
        '23/05/2010: ZaMa - El usuario pierde la pertenencia del npc.
        '13/07/2010: ZaMa - Optimizaciones de logica en la seleccion de pretoriano, y el posible cambio de alencion del usuario.
        '********************************************************

        Dim MiNPC As Npc

        Dim A     As Long
        
100     MiNPC = Npclist(NpcIndex)

        Dim EraCriminal     As Boolean

        Dim PretorianoIndex As Integer
   
        ' @ Reset BOT data
        If MiNPC.BotIndex > 0 Then
            BotIntelligence(MiNPC.BotIndex).Active = False
            MiNPC.BotIndex = 0

        End If

        ' Es pretoriano?
102     If MiNPC.NPCtype = eNPCType.Pretoriano Then
104         Call ClanPretoriano(MiNPC.ClanIndex).MuerePretoriano(NpcIndex)

        End If
        

    
110     If UserIndex > 0 Then
            If MiNPC.CastleIndex > 0 And UserList(UserIndex).GuildIndex > 0 Then
                Castle_Conquist MiNPC.CastleIndex, UserList(UserIndex).GuildIndex
            End If

112         If UserList(UserIndex).flags.SlotEvent > 0 Then

                ' Rey vs Rey
114             If MiNPC.numero = 697 Then FinishCastleMode UserList(UserIndex).flags.SlotEvent, UserList(UserIndex).flags.SlotUserEvent
            
                ' La gran Bestia
116             If MiNPC.numero = 765 Then Call Events_GranBestia_MuereNpc(UserIndex)
    
            End If

        End If
    
        ' Npcs de invocacion
118     If MiNPC.flags.Invocation > 0 Then
120         Invocaciones(MiNPC.flags.Invocation).Activo = 0

        End If
    
        ' Npcs de invasiones
122     If MiNPC.flags.Invasion > 0 Then
124         If Invations(MiNPC.flags.Invasion).Time > 0 Then
126             MiNPC.flags.RespawnTime = 10

            End If

        End If
    
        'Quitamos el npc
128     Call QuitarNPC(NpcIndex)
    
130     If UserIndex > 0 Then ' Lo mato un usuario?

132         With UserList(UserIndex)
        
134             If MiNPC.flags.Snd3 > 0 Then
136                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y, 0))

                End If
            
138             .flags.TargetNPC = 0
140             .flags.TargetNpcTipo = eNPCType.Comun
            
142             If .MascotaIndex Then
144                 Call FollowAmo(.MascotaIndex)

                End If
                
                ' Experiencia de Criaturas restante
146             If MiNPC.flags.ExpCount > 0 Then
148                 If .GroupIndex > 0 Then
150                     Call mGroup.AddExpGroup(UserIndex, MiNPC.flags.ExpCount, MiNPC.GiveGLD)
                    Else
152                     .Stats.Exp = .Stats.Exp + MiNPC.flags.ExpCount

154                     If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                        'Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
156                     Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("Exp +" & CStr(Format(MiNPC.flags.ExpCount, "###,###,###")), d_Exp, 3000, 0))

                    End If

158                 MiNPC.flags.ExpCount = 0

                Else

                    If .GroupIndex > 0 Then
                        Call mGroup.AddExpGroup(UserIndex, 0, MiNPC.GiveGLD)

                    End If

                End If
                
                ' Experiencia de Clan restante
                If MiNPC.flags.ExpGuildCount > 0 Then
                    If UserList(UserIndex).GuildIndex > 0 Then
                        Call Guilds_AddExp(UserIndex, MiNPC.flags.ExpGuildCount)

                    End If
                    
                    MiNPC.flags.ExpGuildCount = 0

                End If
                
                ' Criaturas que dan recursos (leña,fragmentos,minerales,pecesitoh)
                If MiNPC.flags.ResourceCount > 0 Then

                    Dim Obj As Obj

                    Obj.ObjIndex = MiNPC.GiveResource.ObjIndex
                    Obj.Amount = MiNPC.flags.ResourceCount
                        
                    Call MeterItemEnInventario(UserIndex, Obj)
                    
                    MiNPC.flags.ResourceCount = 0

                End If
                
160             If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1

162             Call CheckUserLevel(UserIndex)
            
164             If NpcIndex = .flags.ParalizedByNpcIndex Then
166                 Call RemoveParalisis(UserIndex)

                End If
            
            End With
            
190         If MiNPC.MaestroUser = 0 Then
                'Tiramos el inventario
192             Call NPC_TIRAR_ITEMS(UserIndex, MiNPC, MiNPC.NPCtype = eNPCType.Pretoriano)
        
194             If MiNPC.flags.RespawnTime Then
196                 If MiNPC.flags.Respawn = 0 Then
198                     If Not General.Respawn_Npc_Free(MiNPC.numero, MiNPC.Pos.Map, MiNPC.flags.RespawnTime, MiNPC.CastleIndex, MiNPC.Orig) Then
200                         Call LogError("Ocurrio un error al respawnear el NPC: " & MiNPC.numero & ".")

                        End If

                    End If

                Else
                    'ReSpawn o no
202                 Call RespawnNpc(MiNPC)

                End If

            End If

            Call WriteConsoleMsg(UserIndex, "Acabaste con " & MiNPC.Name, FontTypeNames.FONTTYPE_INFORED)
        End If ' Userindex > 0
    
        
        '<EhFooter>
        Exit Sub

MuereNpc_Err:
        LogError Err.description & vbCrLf & "in MuereNpc " & "at line " & Erl

        '</EhFooter>
End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetNpcFlags_Err
        '</EhHeader>

        'Clear the npc's flags
    
    
100     With Npclist(NpcIndex).flags
              .NpcIdle = False
102         .KeepHeading = 0
104         .RespawnTime = 0
106         .Invasion = 0
108         .Invocation = 0
110         .TeamEvent = 0
112         .InscribedPrevio = 0
114         .SlotEvent = 0
116         .AfectaParalisis = 0
118         .AguaValida = 0
120         .AttackedBy = vbNullString
122         .AttackedByInteger = 0
124         .AttackedFirstBy = vbNullString
126         .BackUp = 0
128         .Bendicion = 0
130         .Domable = 0
132         .Envenenado = 0
134         .Faccion = 0
136         .Follow = False
138         .AtacaDoble = 0
140         .LanzaSpells = 0
142         .Invisible = 0
144         .Maldicion = 0
146         .OldHostil = 0
148         .OldMovement = 0
150         .Paralizado = 0
152         .Inmovilizado = 0
154         .Respawn = 0
156         .RespawnOrigPos = 0
158         .RespawnOrigPosRandom = 0
160         .Snd1 = 0
162         .Snd2 = 0
164         .Snd3 = 0
166         .TierraInvalida = 0
167         .AtacaUsuarios = True
168         .AtacaNPCs = True
169         .AIAlineacion = e_Alineacion.ninguna
        End With

        '<EhFooter>
        Exit Sub

ResetNpcFlags_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ResetNpcFlags " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetNpcCounters_Err
        '</EhHeader>

100     With Npclist(NpcIndex).Contadores
102         .Paralisis = 0
104         .TiempoExistencia = 0
106         .Attack = 0
108         .Descanso = 0
110         .Incinerado = 0
              .UseItem = 0
112         .MovimientoConstante = 0
114         .Velocity = 0
              .RuidoPocion = 0
        End With

        '<EhFooter>
        Exit Sub

ResetNpcCounters_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ResetNpcCounters " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetNpcCharInfo_Err
        '</EhHeader>

100     With Npclist(NpcIndex).Char
102         .Body = 0
104         .CascoAnim = 0
106         .charindex = 0
108         .FX = 0
110         .Head = 0
112         .Heading = 0
114         .loops = 0
116         .ShieldAnim = 0
118         .WeaponAnim = 0
        End With

        '<EhFooter>
        Exit Sub

ResetNpcCharInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ResetNpcCharInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetNpcCriatures_Err
        '</EhHeader>

        Dim j As Long
    
100     With Npclist(NpcIndex)

102         For j = 1 To .NroCriaturas
104             .Criaturas(j).NpcIndex = 0
106             .Criaturas(j).NpcName = vbNullString
108         Next j
        
110         .NroCriaturas = 0
        End With

        '<EhFooter>
        Exit Sub

ResetNpcCriatures_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ResetNpcCriatures " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetExpresiones_Err
        '</EhHeader>

        Dim j As Long
    
100     With Npclist(NpcIndex)

102         For j = 1 To .NroExpresiones
104             .Expresiones(j) = vbNullString
106         Next j
        
108         .NroExpresiones = 0
        End With

        '<EhFooter>
        Exit Sub

ResetExpresiones_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ResetExpresiones " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '22/05/2010: ZaMa - Ahora se resetea el dueño del npc también.
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetNpcMainInfo_Err
        '</EhHeader>

100     With Npclist(NpcIndex)
102         .Attackable = 0
104         .Comercia = 0
106         .GiveEXP = 0
108         .GiveResource.ObjIndex = 0
110         .GiveResource.Amount = 0
112         .RequiredWeapon = 0
114         .AntiMagia = 0
116         .GiveGLD = 0
118         .Hostile = 0
120         .InvReSpawn = 0
122         .QuestNumber = 0
        
124         If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
126         If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
128         If .Owner > 0 Then Call PerdioNpc(.Owner)
        
130         .MaestroUser = 0
132         .MaestroNpc = 0

134         .Owner = 0
              .CaminataActual = 0
136         .Mascotas = 0
138         .Movement = 0
140         .Name = vbNullString
142         .NPCtype = 0
144         .numero = 0
146         .Orig.Map = 0
148         .Orig.X = 0
150         .Orig.Y = 0
152         .PoderAtaque = 0
154         .PoderEvasion = 0
156         .Pos.Map = 0
158         .Pos.X = 0
160         .Pos.Y = 0
162         .SkillDomar = 0
164         .Target = 0
166         .TargetNPC = 0
168         .TipoItems = 0
170         .Veneno = 0
172         .Desc = vbNullString
        
174         .MenuIndex = 0
        
176         .ClanIndex = 0
        
            Dim j As Long

178         For j = 1 To .NroSpells
180             .Spells(j) = 0
182         Next j

        End With
    
184     Call ResetNpcCharInfo(NpcIndex)
186     Call ResetNpcCriatures(NpcIndex)
188     Call ResetExpresiones(NpcIndex)
        '<EhFooter>
        Exit Sub

ResetNpcMainInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ResetNpcMainInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer, _
                     Optional ByVal RespawnTime As Boolean = False)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Now npcs lose their owner
    '***************************************************
    On Error GoTo Errhandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)
        End If
        
        
        .Action = 0
    End With
    
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then

        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1

            If LastNPC < 1 Then Exit Do
        Loop

    End If
      
    If NumNpcs <> 0 Then
        NumNpcs = NumNpcs - 1
    End If

    Exit Sub

Errhandler:
    Call LogError("Error en QuitarNPC")
End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 18/11/2009
    'Kills a pet
    '***************************************************
    On Error GoTo Errhandler

    Dim i        As Integer

    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        If .MascotaIndex Then
            .MascotaIndex = 0
            Call QuitarNPC(NpcIndex)
        End If

    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en QuitarPet. Error: " & Err.number & " Desc: " & Err.description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, _
                                  Optional PuedeAgua As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1
    End If
    
End Function

Public Function CrearNPC(NroNPC As Integer, _
                         mapa As Integer, _
                         OrigPos As WorldPos, _
                         Optional ByVal CustomHead As Integer, _
                         Optional ByVal ForcePos As Boolean = False) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo Errhandler

    'Crea un NPC del tipo NRONPC

    Dim Pos            As WorldPos

    Dim newpos         As WorldPos

    Dim altpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim Iteraciones    As Long

    Dim PuedeAgua      As Boolean

    Dim PuedeTierra    As Boolean
    
    Dim tmpPos         As Long

    Dim nextPos        As Long

    Dim prevPos        As Long

    Dim TipoPos        As Byte
    
    Dim FirstValidPos  As Long
    
    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    nIndex = OpenNPC(NroNPC, LeerNPCs) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Function
    
    ' Cabeza customizada
    If CustomHead <> 0 Then Npclist(nIndex).Char.Head = CustomHead
    
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    'If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
    'Necesita ser respawned en un lugar especifico
        
    If ((Npclist(nIndex).flags.RespawnOrigPos > 0 And Not Npclist(nIndex).flags.RespawnOrigPosRandom > 0) Or ForcePos = True) And InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Npclist(nIndex).Orig.Map = OrigPos.Map
        Npclist(nIndex).Orig.X = OrigPos.X
        Npclist(nIndex).Orig.Y = OrigPos.Y
        Npclist(nIndex).Pos = Npclist(nIndex).Orig
       
    Else
        
        Pos.Map = mapa 'mapa
        altpos.Map = mapa
        
        If PuedeAgua = True Then
            If PuedeTierra = True Then
                TipoPos = RandomNumber(0, 1)
            Else
                TipoPos = 1

            End If

        Else
            TipoPos = 0

        End If
        
        If UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) = 0 Then
            If TipoPos = 1 Then
                TipoPos = 0
            Else
                TipoPos = 1

            End If

        End If
        
        tmpPos = RandomNumber(1, UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos))
        
        nextPos = tmpPos
        prevPos = tmpPos
        
        Do While Not PosicionValida
                    
            ' Posición random
            If Npclist(nIndex).flags.RespawnOrigPosRandom > 0 Then
                Pos.X = RandomNumber(OrigPos.X - Npclist(nIndex).flags.RespawnOrigPosRandom, OrigPos.X + Npclist(nIndex).flags.RespawnOrigPosRandom)
                Pos.Y = RandomNumber(OrigPos.Y - Npclist(nIndex).flags.RespawnOrigPosRandom, OrigPos.Y + Npclist(nIndex).flags.RespawnOrigPosRandom)
            Else
                Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(tmpPos).X
                Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(tmpPos).Y
            End If
        
            
            
            If FirstValidPos = 0 Then FirstValidPos = tmpPos
            
            If LegalPosNPC(Pos.Map, Pos.X, Pos.Y, PuedeAgua, ForcePos) And TestSpawnTrigger(Pos, PuedeAgua) Then
                
                If Not HayPCarea(Pos) Then

                    With Npclist(nIndex)
                        .Pos.Map = Pos.Map
                        .Pos.X = Pos.X
                        .Pos.Y = Pos.Y
                        .Orig = .Pos

                    End With
                    
                    PosicionValida = True

                End If

            End If
            
            If PosicionValida = False Then
                If tmpPos < nextPos Then
                    If nextPos < UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) Then
                        nextPos = nextPos + 1
                        tmpPos = nextPos
                    Else

                        If prevPos > 1 Then
                            prevPos = prevPos - 1
                            tmpPos = prevPos
                        Else

                            If FirstValidPos > 0 Then

                                With Npclist(nIndex)
                                    .Pos.Map = Pos.Map
                                    .Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).X
                                    .Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).Y
                                    .Orig = .Pos

                                End With
                                
                                PosicionValida = True
                            Else
                                Exit Function

                            End If

                        End If

                    End If

                Else

                    If prevPos > 1 Then
                        prevPos = prevPos - 1
                        tmpPos = prevPos
                    Else

                        If nextPos < UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) Then
                            nextPos = nextPos + 1
                            tmpPos = nextPos
                        Else

                            If FirstValidPos > 0 Then

                                With Npclist(nIndex)
                                    .Pos.Map = Pos.Map
                                    .Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).X
                                    .Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).Y
                                    .Orig = .Pos

                                End With
                                
                                PosicionValida = True
                            Else
                                Exit Function

                            End If

                        End If

                    End If

                End If

            End If

        Loop
        
        'asignamos las nuevas coordenas
        Map = Pos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
        
        If Npclist(nIndex).flags.RespawnOrigPosRandom > 0 Then
            Npclist(nIndex).Orig.Map = Map
            Npclist(nIndex).Orig.X = X
            Npclist(nIndex).Orig.Y = Y
            Npclist(nIndex).Pos = Npclist(nIndex).Orig
        End If
        
        '  Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("map: " & Map & " x: " & X & " y:" & Y, FontTypeNames.FONTTYPE_INFO))
    End If
            
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    CrearNPC = nIndex
    Exit Function
  
Errhandler:
    Call LogError("Error" & Err.number & "(" & Err.description & ") en Function CrearNPC de MODULO_NPCs.bas")

End Function

Public Sub MakeNPCChar(ByVal toMap As Boolean, _
                       ByVal sndIndex As Integer, _
                       ByVal NpcIndex As Integer, _
                       ByVal Map As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo MakeNPCChar_Err
        '</EhHeader>
    
        Dim charindex As Integer
        Dim ValidInvi As Boolean
        Dim Name      As String

        Dim Color     As eNickColor
    
100     If Npclist(NpcIndex).Char.charindex = 0 Then
102         charindex = NextOpenCharIndex
104         Npclist(NpcIndex).Char.charindex = charindex
106         CharList(charindex) = NpcIndex
        End If
    
108     MapData(Map, X, Y).NpcIndex = NpcIndex
    
110     If isNPCResucitador(NpcIndex) Then
112         Call Extra.SetAreaResuTheNpc(NpcIndex)
        End If
    
114     With Npclist(NpcIndex)

            ' Castillo: Pretorianos del clan
118         If .Hostile = 0 Then
120             Name = Npclist(NpcIndex).Name
            End If
        
122         Color = eNickColor.ieCastleGuild
  
124         If Not toMap Then
126             Call WriteCharacterCreate(sndIndex, .Char.Body, .Char.BodyAttack, .Char.Head, .Char.Heading, .Char.charindex, X, Y, _
                                        .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, Name, Color, 0, .Char.AuraIndex, .Char.speeding, .flags.NpcIdle, .numero)

            Else
128             Call ModAreas.CreateEntity(NpcIndex, ENTITY_TYPE_NPC, .Pos, .SizeWidth, .SizeWidth)
            End If
        End With
    
        '<EhFooter>
        Exit Sub

MakeNPCChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.MakeNPCChar " & _
               "at line " & Erl & " IN MAP: " & Map & " " & X & " " & Y & " name:" & Npclist(NpcIndex).Name & "."
        
        '</EhFooter>
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, _
                         ByVal Body As Integer, _
                         ByVal Head As Integer, _
                         ByVal Heading As eHeading)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ChangeNPCChar_Err
        '</EhHeader>

100     If NpcIndex > 0 Then

102         With Npclist(NpcIndex).Char
104             .Body = Body
106             .Head = Head
108             .Heading = Heading
            
110             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, .BodyAttack, Head, Heading, .charindex, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, .AuraIndex, False, Npclist(NpcIndex).flags.NpcIdle, False))
            End With

        End If

        '<EhFooter>
        Exit Sub

ChangeNPCChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ChangeNPCChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EraseNPCChar_Err
        '</EhHeader>

100     If Npclist(NpcIndex).Char.charindex <> 0 Then CharList(Npclist(NpcIndex).Char.charindex) = 0

102     If Npclist(NpcIndex).Char.charindex = LastChar Then
104         Do Until CharList(LastChar) > 0
106             LastChar = LastChar - 1
108             If LastChar <= 1 Then Exit Do
            Loop
        End If
    
        'Actualizamos el area
110     Call ModAreas.DeleteEntity(NpcIndex, ENTITY_TYPE_NPC)

        'Quitamos del mapa
112     MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

        'Update la lista npc
114     Npclist(NpcIndex).Char.charindex = 0

        'update NumChars
116     NumChars = NumChars - 1

        '<EhFooter>
        Exit Sub

EraseNPCChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.EraseNPCChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 06/04/2009
        '06/04/2009: ZaMa - Now npcs can force to change position with dead character
        '01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
        '26/09/2010: ZaMa - Turn sub into function to know if npc has moved or not.
        '***************************************************
        '<EhHeader>
        On Error GoTo MoveNPCChar_Err
        '</EhHeader>


        Dim nPos               As WorldPos

        Dim UserIndex          As Integer

        Dim isZonaOscura       As Boolean

        Dim isZonaOscuraNewPos As Boolean
    
100     With Npclist(NpcIndex)
102         nPos = .Pos
104         Call HeadtoPos(nHeading, nPos)
        
        
            
            ' es una posicion legal
            If LegalPosNPC(nPos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0, .flags.TierraInvalida) Then
            
108             If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
110             If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            
112             isZonaOscura = (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura)
114             isZonaOscuraNewPos = (MapData(nPos.Map, nPos.X, nPos.Y).trigger = eTrigger.zonaOscura)
            
116             UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex

                ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
118             If UserIndex > 0 Then
                
                    ' No se traslada caspers de agua a tierra
120                 If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function

                    ' No se traslada caspers de tierra a agua
122                 If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                
124                 With UserList(UserIndex)
                        ' Actualizamos posicion y mapa
126                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
128                     .Pos.X = Npclist(NpcIndex).Pos.X
130                     .Pos.Y = Npclist(NpcIndex).Pos.Y
132                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                        ' Si es un admin invisible, no se avisa a los demas clientes
134                     If Not (.flags.AdminInvisible = 1) Then
136                         Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, .Pos.X, .Pos.Y))
                    
                            'Los valores de visible o invisible están invertidos porque estos flags son del NpcIndex, por lo tanto si el npc entra, el casper sale y viceversa :P
138                         If isZonaOscura Then
140                             If Not isZonaOscuraNewPos Then
142                                 Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
                                End If

                            Else

144                             If isZonaOscuraNewPos Then
146                                 Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
                                End If
                            End If
                        End If
                    
148                     nHeading = InvertHeading(nHeading)
                    
                        'Forzamos al usuario a moverse
150                     Call WriteForceCharMove(UserIndex, nHeading)
                    
                        'Actualizamos las áreas de ser necesario
152                     Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)
                    End With

                End If
                
                
                
                'Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.X, nPos.Y))
                
                'Update map and user pos
154             MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
156             .Pos = nPos
158             .Char.Heading = nHeading
                  .LastHeading = nHeading
160             MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            
162             If isZonaOscura Then
164                 If Not isZonaOscuraNewPos Then
166                     If (.flags.Invisible = 0) Then
168                         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetInvisible(.Char.charindex, False))
                        End If
                    End If

                Else

170                 If isZonaOscuraNewPos Then
172                     If (.flags.Invisible = 0) Then
174                         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetInvisible(.Char.charindex, True))
                        End If
                    End If
                End If
            
176             Call ModAreas.UpdateEntity(NpcIndex, ENTITY_TYPE_NPC, .Pos)
        
                ' Npc has moved
178             MoveNPCChar = True

            End If

        End With
    
        '<EhFooter>
        Exit Function

MoveNPCChar_Err:
        LogError Err.description & vbCrLf & _
               "in MoveNPCChar " & _
               "at line " & Erl

        '</EhFooter>
End Function

Function NextOpenNPC() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo Errhandler

    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1

        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC

    Exit Function

Errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 10/07/2010
        '10/07/2010: ZaMa - Now npcs can't poison dead users.
        '***************************************************
        '<EhHeader>
        On Error GoTo NpcEnvenenarUser_Err
        '</EhHeader>

        Dim N As Integer
    
100     With UserList(UserIndex)

102         If .flags.Muerto = 1 Then Exit Sub
104         If .flags.Envenenado = 1 Then Exit Sub
        
106         N = RandomNumber(1, 100)

108         If N < 30 Then
110             .flags.Envenenado = 1
112             Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
114             Call WriteUpdateEffect(UserIndex)
            End If

        End With
    
        '<EhFooter>
        Exit Sub

NpcEnvenenarUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.NpcEnvenenarUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, _
                  Pos As WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean) As Integer
        '<EhHeader>
        On Error GoTo SpawnNpc_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 06/15/2008
        '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
        '06/15/2008 -> Optimizé el codigo. (NicoNZ)
        '***************************************************
        Dim newpos         As WorldPos

        Dim altpos         As WorldPos

        Dim nIndex         As Integer

        Dim PosicionValida As Boolean

        Dim PuedeAgua      As Boolean

        Dim PuedeTierra    As Boolean

        Dim Map            As Integer

        Dim X              As Integer

        Dim Y              As Integer

100     nIndex = OpenNPC(NpcIndex, LeerNPCs, Respawn)   'Conseguimos un indice

102     If nIndex > MAXNPCS Then
104         SpawnNpc = 0
            Exit Function
        End If

106     PuedeAgua = Npclist(nIndex).flags.AguaValida
108     PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
110     Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra) 'Nos devuelve la posicion valida mas cercana
112     Call ClosestLegalPos(Pos, altpos, PuedeAgua)
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida

114     If newpos.X <> 0 And newpos.Y <> 0 Then
            'Asignamos las nuevas coordenas solo si son validas
116         Npclist(nIndex).Pos.Map = newpos.Map
118         Npclist(nIndex).Pos.X = newpos.X
120         Npclist(nIndex).Pos.Y = newpos.Y
122         PosicionValida = True
        Else

124         If altpos.X <> 0 And altpos.Y <> 0 Then
126             Npclist(nIndex).Pos.Map = altpos.Map
128             Npclist(nIndex).Pos.X = altpos.X
130             Npclist(nIndex).Pos.Y = altpos.Y
132             PosicionValida = True
            Else
134             PosicionValida = False
            End If
        End If

136     If Not PosicionValida Then
138         Call QuitarNPC(nIndex)
140         SpawnNpc = 0
            Exit Function
        End If
    
142     Npclist(nIndex).Orig.Map = Npclist(nIndex).Pos.Map
144     Npclist(nIndex).Orig.X = Npclist(nIndex).Pos.X
146     Npclist(nIndex).Orig.Y = Npclist(nIndex).Pos.Y

        'asignamos las nuevas coordenas
148     Map = newpos.Map
150     X = Npclist(nIndex).Pos.X
152     Y = Npclist(nIndex).Pos.Y

        'Crea el NPC
154     Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

156     If FX Then
158         Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayEffect(SND_WARP, X, Y))
160         Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.charindex, FXIDs.FXWARP, 0))
        End If

162     SpawnNpc = nIndex

        '<EhFooter>
        Exit Function

SpawnNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.SpawnNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub RespawnNpc(MiNPC As Npc)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo RespawnNpc_Err
        '</EhHeader>

100     If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.numero, MiNPC.Pos.Map, MiNPC.Orig)

        '<EhFooter>
        Exit Sub

RespawnNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.RespawnNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, _
                        ByRef ARCHIVE As clsIniManager, _
                        Optional ByVal Respawn = True) As Integer

        On Error GoTo OpenNPC_Err

        Dim NpcIndex As Integer
        
        Dim Field() As String
        
        Dim Leer     As clsIniManager

        Dim LoopC    As Long

        Dim ln       As String
    
100     Set Leer = LeerNPCs
        
        Dim Cabecera As String
        Cabecera = "NPC" & NpcNumber
        
        ' If requested index is invalid, abort
102     If Not Leer.KeyExists(Cabecera) Then
104         OpenNPC = MAXNPCS + 1

            Exit Function

        End If
    
106     NpcIndex = NextOpenNPC
    
108     If NpcIndex > MAXNPCS Then 'Limite de npcs
110         OpenNPC = NpcIndex

            Exit Function

        End If
    
112     With Npclist(NpcIndex)
            ' News
            
            ' Posición utilizada para:
            ' 1° Posición AFK
114         ln = Leer.GetValue("NPC" & NpcNumber, "POSA")
116         .PosA.Map = val(ReadField(1, ln, Asc("-")))
118         .PosA.X = val(ReadField(2, ln, Asc("-")))
120         .PosA.Y = val(ReadField(3, ln, Asc("-")))
            
            ' 2° Posicion Movimiento
122         ln = Leer.GetValue("NPC" & NpcNumber, "POSB")
124         .PosB.Map = val(ReadField(1, ln, Asc("-")))
126         .PosB.X = val(ReadField(2, ln, Asc("-")))
128         .PosB.Y = val(ReadField(3, ln, Asc("-")))
            
            ' 3° Posicion de Ataque
130         ln = Leer.GetValue("NPC" & NpcNumber, "POSC")
132         .PosC.Map = val(ReadField(1, ln, Asc("-")))
134         .PosC.X = val(ReadField(2, ln, Asc("-")))
136         .PosC.Y = val(ReadField(3, ln, Asc("-")))
            ' End News
138         .numero = NpcNumber
140         .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
142         .Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        
144         .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
146         .flags.OldMovement = .Movement
        
148         .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
150         .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
152         .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
154         .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
156         .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
158         .Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
160         .Char.BodyAttack = val(Leer.GetValue("NPC" & NpcNumber, "BodyAttack"))
162         '.Char.AuraIndex(5) = val(Leer.GetValue("NPC" & NpcNumber, "AuraIndex"))
164         .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
166         .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
            .Char.BodyIdle = val(Leer.GetValue("NPC" & NpcNumber, "BodyIdle"))
            
            If .Char.BodyIdle = 0 Then .Char.BodyIdle = .Char.Body
              
168         .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "WeaponAnim"))
170         .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "ShieldAnim"))
172         .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
        
174         .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
176         .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
178         .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
180         .flags.OldHostil = .Hostile
        
182         .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * MultExp
184         .flags.ExpCount = .GiveEXP
        
            .Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
186         .GiveEXPGuild = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPGuild"))
188         .flags.ExpGuildCount = .GiveEXPGuild
        
            ' Recursos de la Criatura
190         ln = Leer.GetValue("NPC" & NpcNumber, "GiveResource")
192         .GiveResource.ObjIndex = val(ReadField(1, ln, 45))
194         .GiveResource.Amount = val(ReadField(2, ln, 45))
              
196         .flags.ResourceCount = .GiveResource.Amount
198         .RequiredWeapon = val(Leer.GetValue("NPC" & NpcNumber, "RequiredWeapon"))
200         .AntiMagia = val(Leer.GetValue("NPC" & NpcNumber, "AntiMagia"))
        
            ' Necesita un Arma Especifica para Atacar
         
202         .Velocity = val(Leer.GetValue("NPC" & NpcNumber, "Velocity"))

            If .Velocity = 0 Then
216             .Velocity = 380
218             .Char.speeding = frmMain.TIMER_AI.interval / 330
            Else
                  
220             .Char.speeding = 210 / .Velocity

                '
            End If
            
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< PATHFINDING >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
             .pathFindingInfo.RangoVision = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
             If .pathFindingInfo.RangoVision = 0 Then .pathFindingInfo.RangoVision = RANGO_VISION_x
            
             .pathFindingInfo.Inteligencia = val(Leer.GetValue("NPC" & NpcNumber, "Inteligencia"))
             If .pathFindingInfo.Inteligencia = 0 Then .pathFindingInfo.Inteligencia = 10
            
             ReDim .pathFindingInfo.Path(1 To .pathFindingInfo.Inteligencia + RANGO_VISION_x * 3)

206         .IntervalAttack = val(Leer.GetValue("NPC" & NpcNumber, "IntervalAttack"))

208         If .IntervalAttack = 0 Then .IntervalAttack = 1500
        
210         .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
212         .Level = val(Leer.GetValue("NPC" & NpcNumber, "ELV"))
214         .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
            .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD")) * MultGld
            .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
        
            .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
            .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
224         .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
226         .MonturaIndex = val(Leer.GetValue("NPC" & NpcNumber, "MonturaIndex"))
228         .ShowName = val(Leer.GetValue("NPC" & NpcNumber, "ShowName"))
        
230         With .Stats
232             .MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
234             .MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
236             .MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
238             .MinHit = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
240             .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
242             .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
244

            End With
            
            .flags.AIAlineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
            ' Forma de atacar de la criatura
            .PretorianAI = val(Leer.GetValue("NPC" & NpcNumber, "PretorianAI"))
            
            .CastleIndex = val(Leer.GetValue("NPC" & NpcNumber, "CastleIndex"))
            
246         .Quest = val(Leer.GetValue("NPC" & NpcNumber, "Quest"))
        
248         If .Quest > 0 Then
250             ReDim .Quests(1 To .Quest) As Byte
            
252             ln = Leer.GetValue("NPC" & NpcNumber, "Quests")
            
254             For LoopC = 1 To .Quest
256                 .Quests(LoopC) = val(ReadField(LoopC, ln, 45))
258             Next LoopC
            
            End If
        
260         .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))

262         If .Invent.NroItems > 0 Then

264             For LoopC = 1 To .Invent.NroItems
266                 ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
268                 .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
270                 .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    
272             Next LoopC

            End If
        
274         .NroDrops = val(Leer.GetValue("NPC" & NpcNumber, "NRODROPS"))
        
276         If .NroDrops > 0 Then

278             For LoopC = 1 To .NroDrops
280                 ln = Leer.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
282                 .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
284                 .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
286                 .Drop(LoopC).Probability = val(ReadField(3, ln, 45))
                      .Drop(LoopC).ProbNum = val(ReadField(4, ln, 45))
288             Next LoopC

            End If
            
290         .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))

292         If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)

294         For LoopC = 1 To .flags.LanzaSpells
296             .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
298         Next LoopC
        
300         If .NPCtype = eNPCType.Entrenador Then
302             .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
304             ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador

306             For LoopC = 1 To .NroCriaturas
308                 .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
310                 .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
312             Next LoopC

            End If
        
314         With .flags
316             .NPCActive = True
            
318             If Respawn Then
320                 .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
                Else
322                 .Respawn = 1

                End If
            
324             .RespawnTime = val(Leer.GetValue("NPC" & NpcNumber, "RespawnTime"))

326             .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
                .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
330             .RespawnOrigPosRandom = val(Leer.GetValue("NPC" & NpcNumber, "OrigPosRandom"))
332             .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
                        
334             .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
336             .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
338             .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

            End With
        
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
340         .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))

342         If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String

344         For LoopC = 1 To .NroExpresiones
346             .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
348         Next LoopC

            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
            ' Menu desplegable p/npc
350         Select Case .NPCtype

                Case eNPCType.Banquero
352                 .MenuIndex = eMenues.ieBanquero
                
354             Case eNPCType.Entrenador
356                 .MenuIndex = eMenues.ieEntrenador
                
358             Case eNPCType.Gobernador
360                 .MenuIndex = eMenues.ieGobernador
                
362             Case eNPCType.Noble
364                 .MenuIndex = eMenues.ieEnlistadorFaccion
                
366             Case eNPCType.ResucitadorNewbie, eNPCType.Revividor
368                 .MenuIndex = eMenues.ieSacerdote
                
370             Case eNPCType.Timbero
372                 .MenuIndex = eMenues.ieApostador
                
374             Case Else

376                 If .flags.Domable <> 0 Then
378                     .MenuIndex = eMenues.ieNpcDomable

                    End If

            End Select
        
380         If .Comercia = 1 Then .MenuIndex = eMenues.ieComerciante
        
            'Tipo de items con los que comercia
382         .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
384         .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
386         .SizeWidth = CByte(val(Leer.GetValue("NPC" & NpcNumber, "SizeWidth")))
388         .SizeHeight = CByte(val(Leer.GetValue("NPC" & NpcNumber, "SizeHeight")))
                
390         If .SizeWidth = 0 Then .SizeWidth = ModAreas.DEFAULT_ENTITY_WIDTH
392         If .SizeHeight = 0 Then .SizeHeight = ModAreas.DEFAULT_ENTITY_HEIGHT
        
394         .EventIndex = CByte(val(Leer.GetValue("NPC" & NpcNumber, "EventIndex")))
            
            ' Por defecto la animación es idle
            If NumUsers > 0 Then
                Call AnimacionIdle(NpcIndex, True)

            End If
            
            ' Si el tipo de movimiento es Caminata
426         If .Movement = Caminata Then
                ' Leemos la cantidad de indicaciones
                Dim cant As Byte
428             cant = val(Leer.GetValue("NPC" & NpcNumber, "CaminataLen"))
                ' Prevengo NPCs rotos
430             If cant = 0 Then
432                 .Movement = Estatico
                Else
                    ' Redimenciono el array
434                 ReDim .Caminata(1 To cant)
                    
                    ' Leo todas las indicaciones
436                 For LoopC = 1 To cant
438                     Field = Split(Leer.GetValue("NPC" & NpcNumber, "Caminata" & LoopC), ":")
    
440                     .Caminata(LoopC).offset.X = val(Field(0))
442                     .Caminata(LoopC).offset.Y = val(Field(1))
444                     .Caminata(LoopC).Espera = val(Field(2))
                    Next
                    
446                 .CaminataActual = 1
                End If
            End If


            If .NroDrops Then
                .TempDrops = NPC_LISTAR_ITEMS(NpcIndex)
            End If
        End With
        
        
        
        
        'Update contadores de NPCs
396     If NpcIndex > LastNPC Then LastNPC = NpcIndex
398     NumNpcs = NumNpcs + 1
    
        'Devuelve el nuevo Indice
400     OpenNPC = NpcIndex
        '<EhFooter>
        Exit Function

OpenNPC_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.OpenNPC " & "at line " & Erl
        
        '</EhFooter>
End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
        
    On Error GoTo 0
        
    With Npclist(NpcIndex)
    
        If .flags.Follow Then
        
            .flags.AttackedBy = vbNullString
            .Target = 0
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
   
        Else
        
            .flags.AttackedBy = UserName
            .Target = NameIndex(UserName)
            .flags.Follow = True
            .Movement = TipoAI.NpcDefensa
            .Hostile = 0

        End If
    
    End With
        
    Exit Sub

DoFollow_Err:
        
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo FollowAmo_Err
        '</EhHeader>

100     With Npclist(NpcIndex)
102         .flags.Follow = True
104         .Movement = TipoAI.SigueAmo
106         .Hostile = 0
108         .Target = 0
110         .TargetNPC = 0
        End With

        '<EhFooter>
        Exit Sub

FollowAmo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.FollowAmo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ValidarPermanenciaNpc(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        'Chequea si el npc continua perteneciendo a algún usuario
        '***************************************************
        '<EhHeader>
        On Error GoTo ValidarPermanenciaNpc_Err
        '</EhHeader>

100     With Npclist(NpcIndex)

102         If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)
        End With

        '<EhFooter>
        Exit Sub

ValidarPermanenciaNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.ValidarPermanenciaNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Sub AnimacionIdle(ByVal NpcIndex As Integer, ByVal Show As Boolean)
    
        On Error GoTo Handler
    
100     With Npclist(NpcIndex)
    
102         If .Char.BodyIdle = 0 Then Exit Sub
        
104         If .flags.NpcIdle = Show Then Exit Sub

106         .flags.NpcIdle = Show
        
108         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, .Char.Heading)
        
        End With
    
        Exit Sub
Handler:
End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado. Se usa para mover NPCs del camino de otro char.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveNpcToSide(ByVal NpcIndex As Integer, ByVal Heading As eHeading)

        On Error GoTo Handler

100     With Npclist(NpcIndex)

            ' Elegimos un lado al azar
            Dim r As Integer
102         r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

            ' Roto el heading original hacia ese lado
104         Heading = Rotate_Heading(Heading, r)

            ' Intento moverlo para ese lado
106         If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
            ' Si falló, intento moverlo para el lado opuesto
108         Heading = InvertHeading(Heading)
110         If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
            ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
            Dim NuevaPos As WorldPos
112         Call ClosestLegalPos(.Pos, NuevaPos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
114         Call WarpNpcChar(NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

        End With

        Exit Sub
    
Handler:

End Sub

Sub WarpNpcChar(ByVal NpcIndex As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
        '<EhHeader>
        On Error GoTo WarpNpcChar_Err
        '</EhHeader>

        Dim NuevaPos                    As WorldPos
        Dim FuturePos                   As WorldPos

100     Call EraseNPCChar(NpcIndex)

102     FuturePos.Map = Map
104     FuturePos.X = X
106     FuturePos.Y = Y
108     Call ClosestLegalPos(FuturePos, NuevaPos, True, True)

110     If NuevaPos.Map = 0 Or NuevaPos.X = 0 Or NuevaPos.Y = 0 Then
112         Debug.Print "Error al tepear NPC"
114         Call QuitarNPC(NpcIndex)
        Else
116         Npclist(NpcIndex).Pos = NuevaPos
118         Call MakeNPCChar(True, 0, NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

120         If FX Then                                    'FX
122             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(SND_WARP, NuevaPos.X, NuevaPos.Y))
124             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, FXIDs.FXWARP, 0))
            End If

        End If

        '<EhFooter>
        Exit Sub

WarpNpcChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.NPCs.WarpNpcChar " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

