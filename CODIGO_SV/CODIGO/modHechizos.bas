Attribute VB_Name = "modHechizos"
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

Public Const SUPERANILLO       As Integer = 700
Public Sub ChangeSlotSpell(ByVal UserIndex As Integer, _
                            ByVal SlotOld As Byte, _
                            ByVal SlotNew As Byte)
        '<EhHeader>
        On Error GoTo ChangeSlotSpell_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
            Dim TempHechizo As Integer
        
102         If SlotOld <= 0 Or SlotOld > MAXUSERHECHIZOS Then
104             Call Logs_Security(eSecurity, eAntiHack, "El personaje " & UserList(UserIndex).Name & " ha intentado hackear el ChangeSlotSpell")
                Exit Sub
            End If
        
106         If SlotNew <= 0 Or SlotNew > MAXUSERHECHIZOS Then
108             Call Logs_Security(eSecurity, eAntiHack, "El personaje " & UserList(UserIndex).Name & " ha intentado hackear el ChangeSlotSpell")
                Exit Sub
            End If
        
        
110         TempHechizo = .Stats.UserHechizos(SlotOld)
112         .Stats.UserHechizos(SlotOld) = .Stats.UserHechizos(SlotNew)
114         .Stats.UserHechizos(SlotNew) = TempHechizo
        End With

        '<EhFooter>
        Exit Sub

ChangeSlotSpell_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.ChangeSlotSpell " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub NpcLanzaSpellSobreTerreno(ByVal NpcIndex As Integer, _
                              ByVal Map As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer)
        '<EhHeader>
        On Error GoTo NpcLanzaSpellSobreTerreno_Err
        '</EhHeader>
                                  
100     If Not Intervalo_CriatureAttack(NpcIndex) Then Exit Sub
102     If Npclist(NpcIndex).flags.LanzaSpells = 0 Then Exit Sub
    
        Dim TempX As Integer
        Dim TempY As Integer
        Dim UserIndex As Integer
        Dim SpellIndex As Integer
        Dim Random As Integer
    
        Dim FxBool As Boolean
104     Random = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
106     SpellIndex = Npclist(NpcIndex).Spells(Random)
    
108     With Hechizos(SpellIndex)
110         For TempX = X - .TileRange To X + .TileRange
112             For TempY = Y - .TileRange To Y + .TileRange
    
114                 If InMapBounds(Map, TempX, TempY) Then
116                     UserIndex = MapData(Map, TempX, TempY).UserIndex
                    
118                     If UserIndex > 0 Then
                            ' ¡¡Agregan HP!!
120                         If .SubeHP = 1 Then
                        
                                ' Update HP
122                             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + RandomNumber(.MinHp, .MaxHp)
124                             If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                            
                            ' ¡¡Quitan HP!!
126                         ElseIf .SubeHP = 2 Then
                            
                            End If
                        End If
                    
128                     If RandomNumber(1, 10) <= 2 And Not FxBool Then
130                         FxBool = True
132                         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFXMap(TempX, TempY, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                        End If
                    End If
                
    
134             Next TempY
136         Next TempX
        
        
138         If Not FxBool Then
140             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFXMap(X, Y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            End If

            ' Spell Wav
142         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, X, Y))

     
            ' Spell Words
144         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, Npclist(NpcIndex).Char.charindex, vbCyan))
        
    
        End With
    

                
        '<EhFooter>
        Exit Sub

NpcLanzaSpellSobreTerreno_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.NpcLanzaSpellSobreTerreno " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, _
                           ByVal UserIndex As Integer, _
                           ByVal Spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 11/11/2010
        '13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
        '13/07/2010: ZaMa - Ahora no se contabiliza la muerte de un atacable.
        '21/09/2010: ZaMa - Amplio los tipos de hechizos que pueden lanzar los npcs.
        '21/09/2010: ZaMa - Permito que se ignore el chequeo de visibilidad (pueden atacar a invis u ocultos).
        '11/11/2010: ZaMa - No se envian los efectos del hechizo si no lo castea.
        '***************************************************
        '<EhHeader>
        On Error GoTo NpcLanzaSpellSobreUser_Err
        '</EhHeader>

100     If Not Intervalo_CriatureAttack(NpcIndex) Then Exit Sub
          If Not IntervaloPuedeRecibirAtaqueCriature(UserIndex) Then Exit Sub
          
          If Not EsObjetivoValido(NpcIndex, UserIndex) Then Exit Sub
          
102     With UserList(UserIndex)
    
104         If .flags.Muerto = 1 Then Exit Sub
106         If (.flags.Mimetizado = 1) And (MapInfo(.Pos.Map).Pk) Then Exit Sub ' // NUEVO
108         If Power.UserIndex = UserIndex Then Exit Sub
        
            ' Doesn't consider if the user is hidden/invisible or not.
110         If Not IgnoreVisibilityCheck Then
112             If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub
            End If
        
            ' Si no se peude usar magia en el mapa, no le deja hacerlo.
114         If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

            Dim daño As Integer
    
            ' Heal HP
116         If Hechizos(Spell).SubeHP = 1 Then
        
118             Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
        
120             daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            
122             .Stats.MinHp = .Stats.MinHp + daño

124             If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
            
126             Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
128             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_CurarSpell))
130             Call WriteUpdateUserStats(UserIndex)
        
                ' Damage
132         ElseIf Hechizos(Spell).SubeHP = 2 Then
            
134             If .flags.Privilegios And PlayerType.User Then
            
136                 Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
138                 daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
140                 daño = daño - (daño * .Stats.UserSkills(eSkill.Resistencia) / 2000)
                
142                 If .Invent.CascoEqpObjIndex > 0 Then
144                     daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
                    End If
                
146                 If .Invent.EscudoEqpObjIndex > 0 Then
148                     daño = daño - RandomNumber(ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMax)
                    End If
                
150                 If .Invent.ArmourEqpObjIndex > 0 Then
152                     daño = daño - RandomNumber(ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMax)
                    End If
                
154                 If .Invent.AnilloEqpObjIndex > 0 Then
156                     daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
                    End If
                
158                 daño = daño - (daño * UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
                
160                 If daño < 0 Then daño = 0
            
162                 .Stats.MinHp = .Stats.MinHp - daño
                
164                 Call SubirSkill(UserIndex, eSkill.Resistencia, True)
166                 Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
168                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoNpc))
170                 Call WriteUpdateUserStats(UserIndex)
                
                    'Muere
172                 If .Stats.MinHp < 1 Then
174                     .Stats.MinHp = 0

176                     If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
178                         RestarCriminalidad (UserIndex)
                        End If
                    
                        Dim MasterIndex As Integer

180                     MasterIndex = Npclist(NpcIndex).MaestroUser
                    
                        '[Barrin 1-12-03]
182                     If MasterIndex > 0 Then
                        
                            ' No son frags los muertos atacables
184                         If .flags.AtacablePor <> MasterIndex Then
                                'Store it!
                                ' Call Statistics.StoreFrag(MasterIndex, UserIndex)
                            
186                             Call ContarMuerte(UserIndex, MasterIndex)
                            End If
                        
188                         Call ActStats(UserIndex, MasterIndex)
                        End If

                        '[/Barrin]
                    
190                     Call UserDie(UserIndex)
                    
                    End If
            
                End If
            
            End If
        
            ' Paralisis/Inmobilize
192         If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
            
194             If .flags.Paralizado = 0 Then
                
196                 Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
                
198                 If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
200                     Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                        Exit Sub

                    End If
                
                    Dim Dividido As Byte
                
202                 If Hechizos(Spell).Inmoviliza = 1 Then
204                     .flags.Inmovilizado = 1
                    End If
                  
206                 .flags.Paralizado = 1
                
208                 If .Clase = eClass.Warrior Or .Clase = eClass.Hunter Then
210                     .Counters.Paralisis = Int(IntervaloParalizado / 2)
                    Else
212                     .Counters.Paralisis = IntervaloParalizado
                    End If
                
214                 If .Invent.ReliquiaSlot > 0 Then
216                     If ObjData(.Invent.ReliquiaObjIndex).EffectUser.AfectaParalisis > 0 Then
218                         .Counters.Paralisis = IntervaloParalizado / ObjData(.Invent.ReliquiaObjIndex).EffectUser.AfectaParalisis

220                         If .Counters.Paralisis <= 0 Then .Counters.Paralisis = 0
                        
222                         WriteConsoleMsg UserIndex, "Tu reliquia ha rechazado el efecto de la parálisis a solo " & Int(.Counters.Paralisis / 40) & " segundos.", FontTypeNames.FONTTYPE_INFO
                        End If
                    End If
                
224                 Call WriteParalizeOK(UserIndex)
                
                End If
            
            End If
        
            ' Stupidity
226         If Hechizos(Spell).Estupidez = 1 Then
             
228             If .flags.Estupidez = 0 Then
            
230                 Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
232                 If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
234                     Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                        Exit Sub

                    End If
                  
236                 .flags.Estupidez = 1
238                 .Counters.Ceguera = IntervaloInvisible
                          
240                 Call WriteDumb(UserIndex)
                
                End If
            End If
        
            ' Blind
242         If Hechizos(Spell).Ceguera = 1 Then
             
244             If .flags.Ceguera = 0 Then
            
246                 Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
248                 If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
250                     Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                        Exit Sub

                    End If
                  
252                 .flags.Ceguera = 1
254                 .Counters.Ceguera = IntervaloInvisible
                          
256                 Call WriteBlind(UserIndex)
                
                End If
            End If
        
            ' Remove Invisibility/Hidden
258         If Hechizos(Spell).RemueveInvisibilidadParcial = 1 Then
                 
260             Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
                 
                'Sacamos el efecto de ocultarse
262             If .flags.Oculto = 1 Then
264                 .Counters.TiempoOculto = 0
266                 .flags.Oculto = 0
268                 Call SetInvisible(UserIndex, .Char.charindex, False)
270                 Call WriteConsoleMsg(UserIndex, "¡Has sido detectado!", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
                Else
                    'sino, solo lo "iniciamos" en la sacada de invisibilidad.
272                 Call WriteConsoleMsg(UserIndex, "Comienzas a hacerte visible.", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
274                 .Counters.Invisibilidad = IntervaloInvisible - 1
                End If
        
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

NpcLanzaSpellSobreUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.NpcLanzaSpellSobreUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendSpellEffects(ByVal UserIndex As Integer, _
                             ByVal NpcIndex As Integer, _
                             ByVal Spell As Integer, _
                             ByVal DecirPalabras As Boolean)
        '<EhHeader>
        On Error GoTo SendSpellEffects_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 11/11/2010
        'Sends spell's wav, fx and mgic words to users.
        '***************************************************
100     With UserList(UserIndex)
            ' Spell Wav
102         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Hechizos(Spell).WAV, .Pos.X, .Pos.Y, .Char.charindex))
            
            ' Spell FX
104         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
            ' Spell Words
106         If DecirPalabras Then
                  
108             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y - 1, -2, d_AddMagicWord, Hechizos(Spell).PalabrasMagicas))
                'Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, vbCyan))
            End If

        End With

        '<EhFooter>
        Exit Sub

SendSpellEffects_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.SendSpellEffects " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, _
                                 ByVal TargetNPC As Integer, _
                                 ByVal SpellIndex As Integer, _
                                 Optional ByVal DecirPalabras As Boolean = False)
        '***************************************************
        'Author: Unknown
        'Last Modification: 21/09/2010
        '21/09/2010: ZaMa - Now npcs can cast a wider range of spells.
        '***************************************************
        '<EhHeader>
        On Error GoTo NpcLanzaSpellSobreNpc_Err
        '</EhHeader>

100     If Not Intervalo_CriatureAttack(NpcIndex) Then Exit Sub
    
        Dim Danio As Integer
    
102     With Npclist(TargetNPC)
        
            ' Spell deals damage??
112         If Hechizos(SpellIndex).SubeHP = 2 Then
            
114             Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            
116             If Npclist(NpcIndex).MaestroUser > 0 Then
118                 Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, Danio)
                      Call Quests_AddNpc(Npclist(NpcIndex).MaestroUser, TargetNPC, Danio)
                End If
        
                ' Deal damage
120             .Stats.MinHp = .Stats.MinHp - Danio
            
                'Muere?
122             If .Stats.MinHp < 1 Then
124                 .Stats.MinHp = 0

126                 If Npclist(NpcIndex).MaestroUser > 0 Then
128                     Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                    Else
130                     Call MuereNpc(TargetNPC, 0)
                    End If
                End If
            
                ' Spell recovers health??
132         ElseIf Hechizos(SpellIndex).SubeHP = 1 Then
                If .Stats.MinHp = .Stats.MaxHp Then Exit Sub
                
134             Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                
                
                ' Recovers health
136             .Stats.MinHp = .Stats.MinHp + Danio
            
138             If .Stats.MinHp > .Stats.MaxHp Then
140                 .Stats.MinHp = .Stats.MaxHp
                End If
            
            End If
            
            
        
            ' Spell Adds/Removes poison?
142         If Hechizos(SpellIndex).Envenena = 1 Then
144             .flags.Envenenado = 1
146         ElseIf Hechizos(SpellIndex).CuraVeneno = 1 Then
148             .flags.Envenenado = 0
            End If

            ' Spells Adds/Removes Paralisis/Inmobility?
150         If Hechizos(SpellIndex).Paraliza = 1 Then
152             .flags.Paralizado = 1
154             .flags.Inmovilizado = 1
156             .Contadores.Paralisis = IntervaloParalizado
            
158         ElseIf Hechizos(SpellIndex).Inmoviliza = 1 Then
160             .flags.Inmovilizado = 1
162             .flags.Paralizado = 0
164             .Contadores.Paralisis = IntervaloParalizado
            
166         ElseIf Hechizos(SpellIndex).RemoverParalisis = 1 Then

168             If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
170                 .flags.Paralizado = 0
172                 .flags.Inmovilizado = 0
174                 .Contadores.Paralisis = 0
                End If
            End If
            
            
            
            
                        ' Spell sound and FX
         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y, .Char.charindex))
            
         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
    
            ' Decir las palabras magicas?
         If DecirPalabras Then
             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, Npclist(NpcIndex).Char.charindex, vbCyan))
           End If
    
        End With

        '<EhFooter>
        Exit Sub

NpcLanzaSpellSobreNpc_Err:
        LogError Err.description & vbCrLf & _
               "in NpcLanzaSpellSobreNpc " & _
               "at line " & Erl

        '</EhFooter>
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo TieneHechizo_Err
        '</EhHeader>
    
        Dim j As Integer

100     For j = 1 To MAXUSERHECHIZOS

102         If UserList(UserIndex).Stats.UserHechizos(j) = i Then
104             TieneHechizo = True

                Exit Function

            End If

        Next


        '<EhFooter>
        Exit Function

TieneHechizo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.TieneHechizo " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, _
                   ByVal Slot As Integer, _
                   Optional ByVal HechizoIndex As Integer = 0)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo AgregarHechizo_Err
        '</EhHeader>

        Dim hIndex As Integer

        Dim j      As Integer

100     With UserList(UserIndex)

102         If HechizoIndex > 0 Then
104             hIndex = HechizoIndex
            Else
106             hIndex = ObjData(.Invent.Object(Slot).ObjIndex).HechizoIndex
            End If
    
108         If Not TieneHechizo(hIndex, UserIndex) Then

                'Buscamos un slot vacio
110             For j = 1 To MAXUSERHECHIZOS

112                 If .Stats.UserHechizos(j) = 0 Then Exit For
114             Next j
            
116             If .Stats.UserHechizos(j) <> 0 Then
118                 Call WriteConsoleMsg(UserIndex, "No tienes espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
                Else
120                 .Stats.UserHechizos(j) = hIndex
122                 Call UpdateUserHechizos(False, UserIndex, CByte(j))
                    'Quitamos del inv el item
124                 Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
                End If

            Else
126             Call WriteConsoleMsg(UserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

AgregarHechizo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.AgregarHechizo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
            
Sub DecirPalabrasMagicas(ByVal SpellWords As String, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 17/11/2009
    '25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
    '17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
    '11/06/2011: CHOTS - Color de dialogos customizables
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        If .flags.AdminInvisible <> 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(SpellWords, .Char.charindex, 5))
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y - 1, -2, d_AddMagicWord, SpellWords))
            
            ' Si estaba oculto, se vuelve visible
            If .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                    Call SetInvisible(UserIndex, .Char.charindex, False)
                End If
            End If
        End If

    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en DecirPalabrasMagicas. Error: " & Err.number & " - " & Err.description)
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

        '<EhHeader>
        On Error GoTo PuedeLanzar_Err

        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2010
        'Last Modification By: ZaMa
        '06/11/09 - Corregida la bonificación de maná del mimetismo en el druida con flauta mágica equipada.
        '19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
        '12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
        '***************************************************
        Dim DruidManaBonus As Single

100     With UserList(UserIndex)

102         If .flags.Muerto Then
104             Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                Exit Function

            End If
            
106         If Hechizos(HechizoIndex).NeedStaff > 0 Then
108             If .Clase = eClass.Mage Then
110                 If .Invent.WeaponEqpObjIndex > 0 Then
112                     If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
114                         Call WriteConsoleMsg(UserIndex, "No posees un báculo lo suficientemente poderoso para poder lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                            Exit Function

                        End If

                    Else
116                     Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un báculo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                        Exit Function

                    End If

                End If

            End If
            
            If Hechizos(HechizoIndex).LvlMin > 0 Then
                If .Stats.Elv < Hechizos(HechizoIndex).LvlMin Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas ser Nivel " & Hechizos(HechizoIndex).LvlMin & " para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                    Exit Function
            
                End If
        
            End If
            
118         If .Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
120             Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                Exit Function

            End If
        
122         If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
124             If .Genero = eGenero.Hombre Then
126                 Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Else
128                 Call WriteConsoleMsg(UserIndex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                End If

                Exit Function

            End If
        
130         If .Stats.MinMan < Hechizos(HechizoIndex).ManaRequerido Then
132             Call WriteConsoleMsg(UserIndex, "No tienes suficiente maná.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If
            
            If .Stats.MinHp < Hechizos(HechizoIndex).HpRequerido Then
133             Call WriteConsoleMsg(UserIndex, "No tienes suficiente vida.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If
        
        End With
    
134     PuedeLanzar = True
        '<EhFooter>
        Exit Function

PuedeLanzar_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.PuedeLanzar " & "at line " & Erl
        
        '</EhFooter>
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef B As Boolean)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo HechizoTerrenoEstado_Err
        '</EhHeader>

        Dim PosCasteadaX As Integer

        Dim PosCasteadaY As Integer

        Dim PosCasteadaM As Integer

        Dim H            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     With UserList(UserIndex)
102         PosCasteadaX = .flags.TargetX
104         PosCasteadaY = .flags.TargetY
106         PosCasteadaM = .flags.TargetMap
        
108         H = .flags.Hechizo
        
110         If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
112             B = True

114             For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
116                 For TempY = PosCasteadaY - 8 To PosCasteadaY + 8

118                     If InMapBounds(PosCasteadaM, TempX, TempY) Then
120                         If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                                'hay un user
122                             If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
124                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.charindex, Hechizos(H).FXgrh, Hechizos(H).loops))
                                End If
                            End If
                        End If

126                 Next TempY
128             Next TempX
        
130             Call InfoHechizo(UserIndex)
            End If
        
            Dim daño As Long
                            
132         If Hechizos(H).SanacionGlobalNpcs = 1 Then
134             B = True

136             For TempX = PosCasteadaX - 2 To PosCasteadaX + 2
138                 For TempY = PosCasteadaY - 2 To PosCasteadaY + 2

140                     If InMapBounds(PosCasteadaM, TempX, TempY) Then
142                         If MapData(PosCasteadaM, TempX, TempY).NpcIndex > 0 Then

144                             Dim tNpc As Integer: tNpc = MapData(PosCasteadaM, TempX, TempY).NpcIndex
                            
146                             daño = RandomNumber(Hechizos(H).MinHp, Hechizos(H).MaxHp)
148                             Npclist(tNpc).Stats.MinHp = Npclist(tNpc).Stats.MinHp + daño
                                
150                             If Npclist(tNpc).Stats.MinHp > Npclist(tNpc).Stats.MaxHp Then Npclist(tNpc).Stats.MinHp = Npclist(tNpc).Stats.MaxHp
                                    
152                             Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateFX(Npclist(tNpc).Char.charindex, Hechizos(H).FXgrh, Hechizos(H).loops))
154                             Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateDamage(Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y, daño, d_CurarSpell))

                            End If
                        End If

156                 Next TempY
158             Next TempX
        
160             Call InfoHechizo(UserIndex)
            End If
        
162         If Hechizos(H).SanacionGlobal = 1 Then
164             If .GuildIndex = 0 Then Exit Sub
166             B = True

168             For TempX = PosCasteadaX - 2 To PosCasteadaX + 2
170                 For TempY = PosCasteadaY - 2 To PosCasteadaY + 2

172                     If InMapBounds(PosCasteadaM, TempX, TempY) Then
174                         If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

176                             Dim tUser As Integer: tUser = MapData(PosCasteadaM, TempX, TempY).UserIndex
                            
                                ' Curamos a nuestro propio CLAN.
178                             If .GuildIndex = UserList(tUser).GuildIndex Then
180                                 daño = RandomNumber(Hechizos(H).MinHp, Hechizos(H).MaxHp)
182                                 UserList(tUser).Stats.MinHp = UserList(tUser).Stats.MinHp + daño
                                
184                                 If UserList(tUser).Stats.MinHp > UserList(tUser).Stats.MaxHp Then UserList(tUser).Stats.MinHp = UserList(tUser).Stats.MaxHp
                                    
186                                 Call WriteUpdateHP(tUser)
188                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.charindex, Hechizos(H).FXgrh, Hechizos(H).loops))
190                                 Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateDamage(UserList(tUser).Pos.X, UserList(tUser).Pos.Y, daño, d_CurarSpell))
                                
                                    ' Sanación más REMOVER PARÁLISIS.
192                                 If Hechizos(H).RemoverParalisis = 1 Then
194                                     If UserList(tUser).flags.Paralizado = 1 Or UserList(tUser).flags.Inmovilizado = 1 Then
196                                         Call RemoveParalisis(tUser)
                                        End If
                                    End If
                                End If
                            End If
                        End If

198                 Next TempY
200             Next TempX
        
202             Call InfoHechizo(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

HechizoTerrenoEstado_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HechizoTerrenoEstado " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Author: Uknown
    'Last modification: 18/09/2010
    'Sale del sub si no hay una posición valida.
    '18/11/2009: Optimizacion de codigo.
    '18/09/2010: ZaMa - No se permite invocar en mapas con InvocarSinEfecto.
    '***************************************************

    On Error GoTo error

    With UserList(UserIndex)

        Dim mapa As Integer

        mapa = .Pos.Map
    
        'No permitimos se invoquen criaturas en zonas seguras
        If MapInfo(mapa).Pk = False Or MapData(mapa, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
            Call WriteConsoleMsg(UserIndex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        'No permitimos se invoquen criaturas en mapas donde esta prohibido hacerlo
        If MapInfo(mapa).InvocarSinEfecto = 1 Then
            Call WriteConsoleMsg(UserIndex, "Invocar no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.SlotFast > 0 Then
            If RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite invocar criaturas.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Sub
    
            End If
        End If
        
        Dim SlotEvent As Byte

        SlotEvent = .flags.SlotEvent

        If SlotEvent > 0 Then
            If Events(SlotEvent).config(eConfigEvent.eInvocar) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Invocar no está permitido aquí! Retirate de la Zona del Evento si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        If .MascotaIndex Then
            Call QuitarPet(UserIndex, .MascotaIndex)
            Exit Sub

        End If
            
        Dim SpellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer

        Dim targetPos  As WorldPos
        
        Dim Entrenable As Boolean
    
        targetPos.Map = .flags.TargetMap
        targetPos.X = .flags.TargetX
        targetPos.Y = .flags.TargetY
    
        SpellIndex = .flags.Hechizo
        
        If MapData(targetPos.Map, targetPos.X, targetPos.Y).trigger = POSINVALIDA Or MapData(targetPos.Map, targetPos.X, targetPos.Y).TileExit.Map <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Elige una posición válida para realizar la invocación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
         
        If Hechizos(SpellIndex).Warp = 1 Then
            PetIndex = Hechizos(SpellIndex).NumNpc

            ' Warp de Mascota
            Entrenable = True
            
        Else
            PetIndex = Hechizos(SpellIndex).NumNpc
            ' Invocación de fuego fatuo y demas
            
           ' If PetIndex = 791 Then
              '  Entrenable = True
           ' Else
             '   Entrenable = False
           ' End If
            
        End If
        
        NpcIndex = SpawnNpc(PetIndex, targetPos, False, False)
            
        If NpcIndex > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            
            .MascotaIndex = NpcIndex
                
            With Npclist(NpcIndex)
                .MaestroUser = UserIndex
                .Contadores.TiempoExistencia = IntervaloInvocacion
                .GiveGLD = 0
                .MenuIndex = eMenues.iemascota
                .Entrenable = Entrenable

            End With
                
            Call FollowAmo(NpcIndex)
        Else

            Exit Sub

        End If

    End With

    Call InfoHechizo(UserIndex)
    HechizoCasteado = True

    Exit Sub

error:

    With UserList(UserIndex)
        LogError ("[" & Err.number & "] " & Err.description & " por el usuario " & .Name & "(" & UserIndex & ") en (" & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")

    End With

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 18/11/2009
        '18/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        '<EhHeader>
        On Error GoTo HandleHechizoTerreno_Err
        '</EhHeader>
    
        Dim HechizoCasteado As Boolean

        Dim ManaRequerida   As Integer
    
100     Select Case Hechizos(SpellIndex).Tipo

            Case TipoHechizo.uInvocacion
102             Call HechizoInvocacion(UserIndex, HechizoCasteado)
            
104         Case TipoHechizo.uEstado
106             Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)
        End Select

108     If HechizoCasteado Then

110         With UserList(UserIndex)
112             Call SubirSkill(UserIndex, eSkill.Magia, True)
            
114             ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
                ' Bonificaciones en hechizos
116             If .Clase = eClass.Druid Then

                    ' Solo con flauta equipada
118                 If .Invent.MagicObjIndex = ANILLOMAGICO Then
                        ' 30% menos de mana para invocaciones
120                     ManaRequerida = ManaRequerida * 0.7
                    End If
                    
                    
                ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                      ' 25% menos de mana para invocaciones
                      ManaRequerida = ManaRequerida * 0.75
                End If
            
                ' Quito la mana requerida
122             .Stats.MinMan = .Stats.MinMan - ManaRequerida

124             If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
                ' Quito la estamina requerida
126             .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

128             If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
                ' Update user stats
130             Call WriteUpdateUserStats(UserIndex)
            End With

        End If
    
        '<EhFooter>
        Exit Sub

HandleHechizoTerreno_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HandleHechizoTerreno " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2010
        '18/11/2009: ZaMa - Optimizacion de codigo.
        '12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
        '***************************************************
        '<EhHeader>
        On Error GoTo HandleHechizoUsuario_Err
        '</EhHeader>
    
        Dim HechizoCasteado As Boolean

        Dim ManaRequerida   As Integer
    
100     Select Case Hechizos(SpellIndex).Tipo

            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoUsuario(UserIndex, HechizoCasteado)
        
104         Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
106             HechizoCasteado = HechizoPropUsuario(UserIndex)
        End Select

108     If HechizoCasteado Then

110         With UserList(UserIndex)
112             Call SubirSkill(UserIndex, eSkill.Magia, True)
            
114             ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
                ' Bonificaciones para druida
116             If .Clase = eClass.Druid Then

                    ' Solo con flauta magica
118                 If .Invent.MagicObjIndex = ANILLOMAGICO Then
120                     If Hechizos(SpellIndex).Mimetiza = 1 Then
                            ' 50% menos de mana para mimetismo
                            ' ManaRequerida = ManaRequerida * 0.5
                        
122                     ElseIf SpellIndex <> APOCALIPSIS_SPELL_INDEX Or SpellIndex <> DESCARGA_SPELL_INDEX Then
                            ' 10% menos de mana para todo menos apoca y descarga
                            'ManaRequerida = ManaRequerida * 0.9
                        End If
                    End If
                    
                    
                ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                      ' 15% menos de mana  hechizos contra usuarios incluyéndose.
                      ManaRequerida = ManaRequerida * 0.85
                End If
            
                ' Quito la mana requerida
124             .Stats.MinMan = .Stats.MinMan - ManaRequerida

126             If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
                ' Quito la estamina requerida
128             .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

130             If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
                ' Update user stats
132             Call WriteUpdateUserStats(UserIndex)
134             Call WriteUpdateUserStats(.flags.TargetUser)
136             .flags.TargetUser = 0
            
            End With

        End If

        '<EhFooter>
        Exit Sub

HandleHechizoUsuario_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HandleHechizoUsuario " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleHechizoNPC_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2010
        '13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
        '17/11/2009: ZaMa - Optimizacion de codigo.
        '12/01/2010: ZaMa - Bonificacion para druidas de 10% para todos hechizos excepto apoca y descarga.
        '12/01/2010: ZaMa - Los druidas mimetizados con npcs ahora son ignorados.
        '***************************************************
        Dim HechizoCasteado As Boolean

        Dim ManaRequerida   As Long
    
100     With UserList(UserIndex)

102         If Npclist(.flags.TargetNPC).AntiMagia > 0 Then
104             Call WriteConsoleMsg(UserIndex, "¡El efecto de Magia ha sido rechazado!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
            
106         Select Case Hechizos(HechizoIndex).Tipo

                Case TipoHechizo.uEstado
                    ' Afectan estados (por ejem : Envenenamiento)
108                 Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, UserIndex)
                
110             Case TipoHechizo.uPropiedades
                    ' Afectan HP,MANA,STAMINA,ETC
112                 Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, UserIndex, HechizoCasteado)
            End Select
        
114         If HechizoCasteado Then
116             Call SubirSkill(UserIndex, eSkill.Magia, True)
            
118             ManaRequerida = Hechizos(HechizoIndex).ManaRequerido
            
                ' Bonificación para druidas.
120             If .Clase = eClass.Druid Then
                    ' Se mostró como usuario, puede ser atacado por npcs
122                 .flags.Ignorado = False
                
                    ' Solo con flauta equipada
124                 If .Invent.MagicObjIndex = ANILLOMAGICO Then
126                     If Hechizos(HechizoIndex).Mimetiza = 1 Then
                            ' 50% menos de mana para mimetismo
128                         ManaRequerida = ManaRequerida * 0.5
                            ' Será ignorado hasta que pierda el efecto del mimetismo o ataque un npc
130                         .flags.Ignorado = True
                        Else

                            ' 10% menos de mana para hechizos
132                         If HechizoIndex <> APOCALIPSIS_SPELL_INDEX Or HechizoIndex <> DESCARGA_SPELL_INDEX Then
                                ' ManaRequerida = ManaRequerida * 0.9
                            End If
                        End If
                    End If
                    
                ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                      ManaRequerida = ManaRequerida * 0.85
                End If
            
                ' Quito la mana requerida
134             .Stats.MinMan = .Stats.MinMan - ManaRequerida

136             If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
                ' Quito la estamina requerida
138             .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido

140             If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
                ' Update user stats
142             Call WriteUpdateUserStats(UserIndex)
144             .flags.TargetNPC = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleHechizoNPC_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HandleHechizoNPC " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LanzarHechizo(ByVal SpellIndex As Integer, ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo LanzarHechizo_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/16/2010
        '24/01/2007 ZaMa - Optimizacion de codigo.
        '02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
        '***************************************************

100     With UserList(UserIndex)
    
102         If .flags.EnConsulta Then
104             Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás en consulta.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
    
106         If .flags.GmSeguidor > 0 Then

                Dim Temp As Long, TiempoActual As Long

108             TiempoActual = GetTime
110             Temp = TiempoActual - .interval(0).ISpell
                        
112             Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 3, Temp, .flags.MenuCliente)
            
                'If .flags.TargetUser > 0 Then
                'Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 5, "Lanzó hechizo sobre " & UserList(.flags.TargetUser).Name, .flags.MenuCliente)
                'ElseIf .flags.TargetNPC > 0 Then
                'Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 5, "Lanzó hechizo sobre " & Npclist(.flags.TargetNPC).Name, .flags.MenuCliente)
                'End If
            
114             .interval(0).ISpell = TiempoActual
            End If
        
            'Chequeamos que no esté desnudo
            'If .flags.Desnudo Then
            'If .Genero = eGenero.Hombre Then
            'Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás desnudo.", FontTypeNames.FONTTYPE_INFO)
            'Else
            'Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás desnuda.", FontTypeNames.FONTTYPE_INFO)
            'End If
            'Exit Sub
            'End If
            
            
    
116         If PuedeLanzar(UserIndex, SpellIndex) Then

118             Select Case Hechizos(SpellIndex).Target

                    Case TargetType.uUsuarios

120                     If .flags.TargetUser > 0 Then
122                         If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
124                             Call HandleHechizoUsuario(UserIndex, SpellIndex)
                            Else
126                             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)
                            End If

                        Else
128                         Call WriteConsoleMsg(UserIndex, "Este hechizo actúa sólo sobre usuarios.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                              'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.flags.TargetX, .flags.TargetY - 1, -1, eDamageType.d_Fallas, "Fallas"))
                        End If
            
130                 Case TargetType.uNPC

132                     If .flags.TargetNPC > 0 Then
134                         If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
136                             Call HandleHechizoNPC(UserIndex, SpellIndex)
                            Else
138                             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)
                            End If

                        Else
140                         Call WriteConsoleMsg(UserIndex, "Este hechizo sólo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                              'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.flags.TargetX, .flags.TargetY - 1, -1, eDamageType.d_Fallas, "Fallas"))
                        End If
            
142                 Case TargetType.uUsuariosYnpc

144                     If .flags.TargetUser > 0 Then
146                         If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
148                             Call HandleHechizoUsuario(UserIndex, SpellIndex)
                            Else
150                             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)
                            End If

152                     ElseIf .flags.TargetNPC > 0 Then

154                         If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
156                             Call HandleHechizoNPC(UserIndex, SpellIndex)
                            Else
158                             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)
                            End If

                        Else
                              'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.flags.TargetX, .flags.TargetY - 1, -1, eDamageType.d_Fallas, "Fallas"))
160                         Call WriteConsoleMsg(UserIndex, "Target inválido.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                        End If
            
162                 Case TargetType.uTerreno
164                     Call HandleHechizoTerreno(UserIndex, SpellIndex)

                      Case TargetType.uArea
                          Call HandleHechizoArea(UserIndex, SpellIndex)
                          
                End Select
        
            End If
    
166         If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
    
168         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
       
        End With

    
        '<EhFooter>
        Exit Sub


LanzarHechizo_Err:
        
        LogError "Error en LanzarHechizo. Error " & Err.number & " : " & Err.description & " Hechizo: " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & "). at line " & Erl
        
        '</EhFooter>
End Sub

Sub HandleHechizoArea(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

        On Error GoTo HandleHechizoArea_Err

        Dim HechizoCasteado As Boolean

        Dim ManaRequerida   As Integer
    
100     Select Case Hechizos(SpellIndex).Tipo

            Case TipoHechizo.uInvocacion
102             'Call HechizoInvocacion(UserIndex, HechizoCasteado)

104         Case TipoHechizo.uPropiedades
                   ' If esnpc then
                        'HechizoCasteado = HechizoPropAreaNPC(UserIndex)
                   ' else
                        HechizoCasteado = HechizoPropAreaUsuario(UserIndex)
                   ' end if
                   
                   
                   
        End Select

108     If HechizoCasteado Then

110         With UserList(UserIndex)
112             Call SubirSkill(UserIndex, eSkill.Magia, True)
            
114             ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
                ' Bonificaciones en hechizos
116             If .Clase = eClass.Druid Then

                    ' Solo con flauta equipada
118                 If .Invent.MagicObjIndex = ANILLOMAGICO Then
                        ' 30% menos de mana para invocaciones
120                     ManaRequerida = ManaRequerida * 0.7
                    End If
                    
                    
                ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                      ' 25% menos de mana para invocaciones
                      ManaRequerida = ManaRequerida * 0.75
                End If
            
                ' Quito la mana requerida
122             .Stats.MinMan = .Stats.MinMan - ManaRequerida

124             If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
                ' Quito la estamina requerida
126             .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

128             If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
                ' Update user stats
130             Call WriteUpdateUserStats(UserIndex)
            End With

        End If
    
        '<EhFooter>
        Exit Sub

HandleHechizoArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HandleHechizoArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 28/04/2010
        'Handles the Spells that afect the Stats of an User
        '24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
        '26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
        '26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
        '02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
        '06/28/2008 NicoNZ - Agregué que se le de valor al flag Inmovilizado.
        '17/11/2008: NicoNZ - Agregado para quitar la penalización de vida en el ring y cambio de ecuacion.
        '13/02/2009: ZaMa - Arreglada ecuacion para quitar vida tras resucitar en rings.
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
        '16/09/2010: ZaMa - Solo se hace invi para los clientes si no esta navegando.
        '***************************************************
        '<EhHeader>
        On Error GoTo HechizoEstadoUsuario_Err

        '</EhHeader>

        Dim HechizoIndex As Integer

        Dim TargetIndex  As Integer

100     With UserList(UserIndex)
102         HechizoIndex = .flags.Hechizo
104         TargetIndex = .flags.TargetUser
    
            ' <-------- Agrega Invisibilidad ---------->
106         If Hechizos(HechizoIndex).Invisibilidad = 1 Then
108             If UserList(TargetIndex).flags.Muerto = 1 Then
110                 Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
112                 HechizoCasteado = False

                    Exit Sub

                End If
        
114             If UserList(TargetIndex).Counters.Saliendo Then
116                 If UserIndex <> TargetIndex Then
118                     Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
120                     HechizoCasteado = False

                        Exit Sub

                    Else
122                     Call WriteConsoleMsg(UserIndex, "¡No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
124                     HechizoCasteado = False

                        Exit Sub

                    End If

                End If
        
                'No usar invi mapas InviSinEfecto
126             If MapInfo(UserList(TargetIndex).Pos.Map).InviSinEfecto > 0 Then
128                 Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
130                 HechizoCasteado = False

                    Exit Sub

                End If
            
132             If .flags.SlotEvent > 0 Then
134                 If Events(.flags.SlotEvent).config(eConfigEvent.eInvisibilidad) = 0 Then
136                     Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí! Retirate de la Zona del Evento si deseas utilizar el hechizo.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
                
                If .flags.SlotFast > 0 Then
                    If RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                        Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub

                    End If
                End If
        
                ' No invi en zona segura
138             If Not MapInfo(.Pos.Map).Pk Then
140                 Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto en zona segura.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
        
142             If UserList(TargetIndex).flags.Mimetizado = 1 Then
144                 Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto estando mimetizado.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
146             If Power.UserIndex = TargetIndex Then
148                 Call WriteConsoleMsg(UserIndex, "El personaje posee un poder superior.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
150             If UserList(TargetIndex).flags.Invisible = 1 Or UserList(TargetIndex).flags.Oculto = 1 Then
152                 Call WriteConsoleMsg(UserIndex, "El personaje ya se encuentra invisible.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                ' Chequea si el status permite ayudar al otro usuario
154             HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)

156             If Not HechizoCasteado Then Exit Sub
        
                'Si sos user, no uses este hechizo con GMS.
158             If .flags.Privilegios And PlayerType.User Then
160                 If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
162                     HechizoCasteado = False

                        Exit Sub

                    End If

                End If
            
164             UserList(TargetIndex).flags.Invisible = 1
            
                ' Solo se hace invi para los clientes si no esta navegando
166             If UserList(TargetIndex).flags.Navegando = 0 Then
168                 Call SetInvisible(TargetIndex, UserList(TargetIndex).Char.charindex, True)
170                 UserList(TargetIndex).Counters.DrawersCount = RandomNumberPower(1, 200)

                End If
        
172             Call InfoHechizo(UserIndex)
            
174             HechizoCasteado = True

            End If
    
            ' <-------- Agrega Mimetismo ---------->
176         If Hechizos(HechizoIndex).Mimetiza = 1 Then
178             If TargetIndex = UserIndex Then Exit Sub
            
180             If UserList(TargetIndex).flags.Muerto = 1 Then

                    Exit Sub

                End If
            
182             If UserList(UserIndex).flags.Navegando = 1 Then

                    Exit Sub

                End If
        
184             If UserList(TargetIndex).flags.Navegando = 1 Then

                    Exit Sub

                End If
        
186             If UserList(TargetIndex).flags.Transform = 1 Then

                    Exit Sub

                End If

192             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
194             If UserList(TargetIndex).flags.Transform = 1 Then

                    Exit Sub

                End If
            
196             If UserList(TargetIndex).flags.TransformVIP = 1 Then

                    Exit Sub

                End If
            
198             If Not MapInfo(.Pos.Map).Pk Then
200                 Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto en zona segura.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
202             If MapInfo(.Pos.Map).MimetismoSinEfecto = 1 Then
204                 Call WriteConsoleMsg(UserIndex, "El mapa no permite el efecto mimetismo.", FontTypeNames.FONTTYPE_INFO)
            
                    Exit Sub

                End If
            
                'Si sos user, no uses este hechizo con GMS.
206             If .flags.Privilegios And PlayerType.User Then
208                 If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then

                        Exit Sub

                    End If

                End If
        
210             If .flags.Mimetizado = 1 Then
212                 Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
        
        
214             If .flags.AdminInvisible = 1 Then Exit Sub

216             If UserList(TargetIndex).flags.Invisible = 1 Or UserList(TargetIndex).flags.Oculto = 1 Then
218                 Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto estando invisible.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
        
        
                'copio el char original al mimetizado
        
220             .CharMimetizado.Body = .Char.Body
222             .CharMimetizado.Head = .Char.Head
224             .CharMimetizado.CascoAnim = .Char.CascoAnim
226             .CharMimetizado.ShieldAnim = .Char.ShieldAnim
228             .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
230             .flags.Mimetizado = 1
232             .flags.Ignorado = True
            
                'ahora pongo local el del enemigo
234             .Char.Body = UserList(TargetIndex).Char.Body
236             .Char.Head = UserList(TargetIndex).Char.Head
238             .Char.CascoAnim = UserList(TargetIndex).Char.CascoAnim
240             .Char.ShieldAnim = UserList(TargetIndex).Char.ShieldAnim
242             .Char.WeaponAnim = UserList(TargetIndex).Char.WeaponAnim
        
244             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
       
246             Call InfoHechizo(UserIndex)
248             HechizoCasteado = True

            End If
    
            ' <-------- Agrega Envenenamiento ---------->
250         If Hechizos(HechizoIndex).Envenena = 1 Then
252             If UserIndex = TargetIndex Then
254                 Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
        
256             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
258             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
260             If UserIndex <> TargetIndex Then
262                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If

264             UserList(TargetIndex).flags.Envenenado = 1
266             Call InfoHechizo(UserIndex)
268             HechizoCasteado = True

            End If
    
            ' <-------- Cura Envenenamiento ---------->
270         If Hechizos(HechizoIndex).CuraVeneno = 1 Then
            
272             If UserList(TargetIndex).flags.Envenenado = 0 Then
274                 Call WriteConsoleMsg(UserIndex, "El personaje no está envenenado.", FontTypeNames.FONTTYPE_INFORED)
276                 HechizoCasteado = False
                    Exit Sub

                End If
            
                'Verificamos que el usuario no este muerto
278             If UserList(TargetIndex).flags.Muerto = 1 Then
280                 Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
282                 HechizoCasteado = False

                    Exit Sub

                End If
        
                ' Chequea si el status permite ayudar al otro usuario
284             HechizoCasteado = CanSupportUser(UserIndex, TargetIndex)

286             If Not HechizoCasteado Then Exit Sub
            
                'Si sos user, no uses este hechizo con GMS.
288             If .flags.Privilegios And PlayerType.User Then
290                 If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then

                        Exit Sub

                    End If

                End If
            
292             UserList(TargetIndex).flags.Envenenado = 0
294             Call WriteUpdateEffect(TargetIndex)
296             Call InfoHechizo(UserIndex)
298             HechizoCasteado = True

            End If
    
            ' <-------- Agrega Maldicion ---------->
300         If Hechizos(HechizoIndex).Maldicion = 1 Then
302             If UserIndex = TargetIndex Then
304                 Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
        
306             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
308             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
310             If UserIndex <> TargetIndex Then
312                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If

314             UserList(TargetIndex).flags.Maldicion = 1
316             Call InfoHechizo(UserIndex)
318             HechizoCasteado = True

            End If
    
            ' <-------- Remueve Maldicion ---------->
320         If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
322             UserList(TargetIndex).flags.Maldicion = 0
324             Call InfoHechizo(UserIndex)
326             HechizoCasteado = True

            End If
    
            ' <-------- Agrega Bendicion ---------->
328         If Hechizos(HechizoIndex).Bendicion = 1 Then
330             UserList(TargetIndex).flags.Bendicion = 1
332             Call InfoHechizo(UserIndex)
334             HechizoCasteado = True

            End If
    
            ' <-------- Agrega Paralisis/Inmobilidad ---------->
336         If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
338             If UserIndex = TargetIndex Then
340                 Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
            
342             If .flags.SlotReto > 0 Then
344                 If Retos(.flags.SlotReto).config(eRetoConfig.eInmovilizar) = 0 Then
346                     Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                        Exit Sub
                
                    End If
            
                End If
            
348             If .flags.SlotEvent > 0 Then
350                 If Events(.flags.SlotEvent).config(eConfigEvent.eUseParalizar) = 0 Then
352                     Call WriteConsoleMsg(UserIndex, "No tienes permitido utilizar esta clase de hechizos en el evento.", FontTypeNames.FONTTYPE_INFORED)
                        Exit Sub

                    End If

                End If
            
354             If UserList(TargetIndex).flags.Paralizado = 0 Then
356                 If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
            
358                 If UserIndex <> TargetIndex Then
                        Call checkHechizosEfectividad(UserIndex, TargetIndex)
360                     Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                    End If
            
362                 Call InfoHechizo(UserIndex)
364                 HechizoCasteado = True

366                 If UserList(TargetIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
368                     Call WriteConsoleMsg(TargetIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
370                     Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
372                     Call FlushBuffer(TargetIndex)

                        Exit Sub

                    End If
                
374                 If Power.UserIndex = TargetIndex Then
376                     Call WriteConsoleMsg(TargetIndex, "¡Te han querido inmovilizar!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
378                     Call WriteConsoleMsg(UserIndex, " ¡Ingenuo! El poder de las medusas es superior al tuyo", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
380                     Call FlushBuffer(TargetIndex)

                        Exit Sub

                    End If
            
382                 If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(TargetIndex).flags.Inmovilizado = 1
384                 UserList(TargetIndex).flags.Paralizado = 1
386                 UserList(TargetIndex).Counters.Paralisis = IIf(.Stats.MaxMan = 0, (IntervaloParalizado / 2), IntervaloParalizado)
            
388                 UserList(TargetIndex).flags.ParalizedByIndex = UserIndex
390                 UserList(TargetIndex).flags.ParalizedBy = UserList(UserIndex).Name
                
                    If UserList(TargetIndex).flags.SlotEvent = 0 Then
                        Call SendData(SendTarget.ToOne, TargetIndex, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " lanzó " & Hechizos(HechizoIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT))
                
                        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageConsoleMsg(Hechizos(HechizoIndex).Nombre & " sobre " & UserList(TargetIndex).Name, FontTypeNames.FONTTYPE_FIGHT))

                    End If
                
392                 Call WriteParalizeOK(TargetIndex)
394                 Call FlushBuffer(TargetIndex)

                End If

            End If
    
            ' <-------- Remueve Paralisis/Inmobilidad ---------->
396         If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
                ' Remueve si esta en ese estado
398             If UserList(TargetIndex).flags.Paralizado = 1 Then
        
                    ' Chequea si el status permite ayudar al otro usuario
400                 HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)

402                 If Not HechizoCasteado Then Exit Sub
                      
404                 Call RemoveParalisis(TargetIndex)
406                 Call InfoHechizo(UserIndex)
                      Call WriteConsoleMsg(TargetIndex, "¡" & .Name & " te ha devuelto la movilidad!", FontTypeNames.FONTTYPE_USERPLATA, eMessageType.Combate)
                      Call WriteConsoleMsg(UserIndex, "¡Has devuelvo la movilidad a " & UserList(TargetIndex).Name & "!", FontTypeNames.FONTTYPE_USERPLATA, eMessageType.Combate)
        
                End If

            End If
    
            ' <-------- Remueve Estupidez (Aturdimiento) ---------->
408         If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
    
                ' Remueve si esta en ese estado
410             If UserList(TargetIndex).flags.Estupidez = 1 Then
        
                    ' Chequea si el status permite ayudar al otro usuario
412                 HechizoCasteado = CanSupportUser(UserIndex, TargetIndex)

414                 If Not HechizoCasteado Then Exit Sub
        
416                 UserList(TargetIndex).flags.Estupidez = 0
            
                    'no need to crypt this
418                 Call WriteDumbNoMore(TargetIndex)
420                 Call FlushBuffer(TargetIndex)
422                 Call InfoHechizo(UserIndex)
        
                End If

            End If
    
            ' <-------- Revive ---------->
424         If Hechizos(HechizoIndex).Revivir = 1 Then
426             If UserList(TargetIndex).flags.Muerto = 1 Then
            
                    'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
428                 If UserList(TargetIndex).flags.SeguroResu Then
430                     Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
432                     HechizoCasteado = False

                        Exit Sub

                    End If
        
                    'No usar resu en mapas con ResuSinEfecto
434                 If MapInfo(UserList(TargetIndex).Pos.Map).ResuSinEfecto > 0 Then
436                     Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
438                     HechizoCasteado = False

                        Exit Sub

                    End If
                
440                 If .flags.SlotReto > 0 Then
442                     If Retos(.flags.SlotReto).config(eRetoConfig.eResucitar) = 0 Then
444                         Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
    
                            Exit Sub
                    
                        End If
                
                    End If
                
446                 If .flags.SlotEvent > 0 Then
448                     If Events(.flags.SlotEvent).config(eConfigEvent.eResu) = 0 Then
450                         Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                            Exit Sub

                        End If

                    End If
                    
                    If .flags.SlotFast > 0 Then
                        If RetoFast(.flags.SlotFast).ConfigVale <> ValeResu And RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                            Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                             Exit Sub

                        End If

                    End If
                    
            
                    'revisamos si necesita vara
452                 If .Clase = eClass.Mage Then
454                     If .Invent.WeaponEqpObjIndex > 0 Then
456                         If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
458                             Call WriteConsoleMsg(UserIndex, "Necesitas un báculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
460                             HechizoCasteado = False

                                Exit Sub

                            End If

                        End If

462                 ElseIf .Clase = eClass.Bard Then

464                     If .Invent.MagicObjIndex <> LAUDMAGICO Then
466                         Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
468                         HechizoCasteado = False

                            Exit Sub

                        End If

470                 ElseIf .Clase = eClass.Druid Then

472                     If .Invent.MagicObjIndex <> ANILLOMAGICO Then
474                         Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
476                         HechizoCasteado = False

                            Exit Sub

                        End If

                    End If
            
                    ' Chequea si el status permite ayudar al otro usuario
478                 HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)

480                 If Not HechizoCasteado Then Exit Sub
    
                    Dim EraCriminal As Boolean

482                 EraCriminal = Escriminal(UserIndex)
            
484                 If Not Escriminal(TargetIndex) Then
486                     If TargetIndex <> UserIndex Then
488                         .Reputacion.NobleRep = .Reputacion.NobleRep + 500

490                         If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
492                         Call WriteConsoleMsg(UserIndex, "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
            
494                 If EraCriminal And Not Escriminal(UserIndex) Then
496                     Call RefreshCharStatus(UserIndex)

                    End If
            
498                 With UserList(TargetIndex)
                        'Pablo Toxic Waste (GD: 29/04/07)
500                     .Stats.MinAGU = 0
502                     .flags.Sed = 1
504                     .Stats.MinHam = 0
506                     .flags.Hambre = 1
508                     Call WriteUpdateHungerAndThirst(TargetIndex)
510                     Call InfoHechizo(UserIndex)
512                     .Stats.MinMan = 0
514                     .Stats.MinSta = 0

                    End With
            
                    'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
516                 If (TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE) Then

                        'Solo saco vida si es User. no quiero que exploten GMs por ahi.
518                     If .flags.Privilegios And PlayerType.User Then
                            If .Clase <> eClass.Cleric Then
520                             .Stats.MinHp = .Stats.MinHp * (1 - (.Stats.Elv) * 0.015)

                            End If

                        End If

                    End If
            
522                 If (.Stats.MinHp <= 0) Then
524                     Call UserDie(UserIndex)
526                     Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
528                     HechizoCasteado = False
                    Else

                        If .Clase <> eClass.Cleric Then
530                         Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)

                        End If

532                     HechizoCasteado = True

                    End If
            
534                 If UserList(TargetIndex).flags.Traveling = 1 Then
536                     Call EndTravel(TargetIndex, True)

                    End If
            
538                 Call RevivirUsuario(TargetIndex)
                Else
540                 HechizoCasteado = False

                End If
    
            End If
    
            ' <-------- Agrega Ceguera ---------->
542         If Hechizos(HechizoIndex).Ceguera = 1 Then
544             If UserIndex = TargetIndex Then
546                 Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
            
548             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
550             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
552             If UserIndex <> TargetIndex Then
554                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If

556             UserList(TargetIndex).flags.Ceguera = 1
558             UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado / 3
    
560             Call WriteBlind(TargetIndex)
562             Call FlushBuffer(TargetIndex)
564             Call InfoHechizo(UserIndex)
566             HechizoCasteado = True

            End If
    
            ' <-------- Agrega Estupidez (Aturdimiento) ---------->
568         If Hechizos(HechizoIndex).Estupidez = 1 Then
570             If UserIndex = TargetIndex Then
572                 Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
            
574             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
576             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
578             If UserIndex <> TargetIndex Then
580                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If

582             If UserList(TargetIndex).flags.Estupidez = 0 Then
584                 UserList(TargetIndex).flags.Estupidez = 1
586                 UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado

                End If

588             Call WriteDumb(TargetIndex)
590             Call FlushBuffer(TargetIndex)
    
592             Call InfoHechizo(UserIndex)
594             HechizoCasteado = True

            End If

        End With

        '<EhFooter>
        Exit Sub

HechizoEstadoUsuario_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoEstadoUsuario " & "at line " & Erl

        '</EhFooter>
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, _
                     ByVal SpellIndex As Integer, _
                     ByRef HechizoCasteado As Boolean, _
                     ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 07/07/2008
        'Handles the Spells that afect the Stats of an NPC
        '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
        'removidos por users de su misma faccion.
        '07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
        '***************************************************
        '<EhHeader>
        On Error GoTo HechizoEstadoNPC_Err
        '</EhHeader>

100     With Npclist(NpcIndex)

102         If Hechizos(SpellIndex).Invisibilidad = 1 Then
104             Call InfoHechizo(UserIndex)
106             .flags.Invisible = 1
108             HechizoCasteado = True
            End If
    
110         If Hechizos(SpellIndex).Envenena = 1 Then
112             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
114                 HechizoCasteado = False

                    Exit Sub

                End If

116             Call NPCAtacado(NpcIndex, UserIndex)
118             Call InfoHechizo(UserIndex)
120             .flags.Envenenado = 1
122             HechizoCasteado = True
            End If
    
124         If Hechizos(SpellIndex).CuraVeneno = 1 Then
126             Call InfoHechizo(UserIndex)
128             .flags.Envenenado = 0
130             HechizoCasteado = True
            End If
    
132         If Hechizos(SpellIndex).Maldicion = 1 Then
134             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
136                 HechizoCasteado = False

                    Exit Sub

                End If

138             Call NPCAtacado(NpcIndex, UserIndex)
140             Call InfoHechizo(UserIndex)
142             .flags.Maldicion = 1
144             HechizoCasteado = True
            End If
    
146         If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
148             Call InfoHechizo(UserIndex)
150             .flags.Maldicion = 0
152             HechizoCasteado = True
            End If
    
154         If Hechizos(SpellIndex).Bendicion = 1 Then
156             Call InfoHechizo(UserIndex)
158             .flags.Bendicion = 1
160             HechizoCasteado = True
            End If
    
162         If Hechizos(SpellIndex).Paraliza = 1 And Hechizos(SpellIndex).Inmoviliza = 1 Then
164             If .flags.AfectaParalisis = 0 Then
166                 If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
168                     HechizoCasteado = False

                        Exit Sub

                    End If

170                 Call NPCAtacado(NpcIndex, UserIndex)
172                 Call InfoHechizo(UserIndex)
174                 .flags.Paralizado = 1
176                 .flags.Inmovilizado = 1
178                 .Contadores.Paralisis = (IntervaloParalizado * 4)
                      Call AnimacionIdle(NpcIndex, False)
180                 HechizoCasteado = True
                Else
182                 Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
184                 HechizoCasteado = False

                    Exit Sub

                End If
            End If
    
186         If Hechizos(SpellIndex).RemoverParalisis = 1 Then
188             If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
190                 If .MaestroUser = UserIndex Then
192                     Call InfoHechizo(UserIndex)
194                     .flags.Paralizado = 0
196                     .Contadores.Paralisis = 0
198                     HechizoCasteado = True
                    Else

200                     If .NPCtype = eNPCType.GuardiaReal Then
202                         If esArmada(UserIndex) Then
204                             Call InfoHechizo(UserIndex)
206                             .flags.Paralizado = 0
208                             .Contadores.Paralisis = 0
210                             HechizoCasteado = True

                                Exit Sub

                            Else
212                             Call WriteConsoleMsg(UserIndex, "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
214                             HechizoCasteado = False

                                Exit Sub

                            End If
                    
216                         Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
218                         HechizoCasteado = False

                            Exit Sub

                        Else

220                         If .NPCtype = eNPCType.GuardiasCaos Then
222                             If esCaos(UserIndex) Then
224                                 Call InfoHechizo(UserIndex)
226                                 .flags.Paralizado = 0
228                                 .Contadores.Paralisis = 0
230                                 HechizoCasteado = True

                                    Exit Sub

                                Else
232                                 Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
234                                 HechizoCasteado = False

                                    Exit Sub

                                End If
                            End If
                        End If
                    End If

                Else
236                 Call WriteConsoleMsg(UserIndex, "Este NPC no está paralizado", FontTypeNames.FONTTYPE_INFO)
238                 HechizoCasteado = False

                    Exit Sub

                End If
            End If
     
240         If Hechizos(SpellIndex).Paraliza = 1 And Hechizos(SpellIndex).Inmoviliza = 0 Then
242             If .flags.AfectaParalisis = 0 Then
244                 If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
246                     HechizoCasteado = False

                        Exit Sub

                    End If

248                 Call NPCAtacado(NpcIndex, UserIndex)
250                 .flags.Inmovilizado = 1
252                 .flags.Paralizado = 0
254                 .Contadores.Paralisis = (IntervaloParalizado * 3)
256                 Call InfoHechizo(UserIndex)
                      Call AnimacionIdle(NpcIndex, True)
258                 HechizoCasteado = True
                Else
260                 Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If

        End With

262     If Hechizos(SpellIndex).Mimetiza = 1 Then

264         With UserList(UserIndex)

266             If .flags.Mimetizado = 1 Then
268                 Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
270             If .flags.Navegando = 1 Then
272                 Call WriteConsoleMsg(UserIndex, "No puedes mimetizarte navegando.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
274             If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
276                 Call WriteConsoleMsg(UserIndex, "No puedes mimetizarte estando invisible.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
            
278             If .flags.Transform = 1 Or .flags.TransformVIP Then
280                 Call WriteConsoleMsg(UserIndex, "No puedes mimetizarte en ese estado.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
            
282             If Not MapInfo(.Pos.Map).Pk Then
284                 Call WriteConsoleMsg(UserIndex, "El hechizo tiene efecto en zonas inseguras", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
        
286             If MapInfo(.Pos.Map).MimetismoSinEfecto = 1 Then
288                 Call WriteConsoleMsg(UserIndex, "El mapa no permite el efecto mimetismo.", FontTypeNames.FONTTYPE_INFO)
            
                    Exit Sub

                End If
                
                
                If Npclist(NpcIndex).Char.Body = 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes tomar la forma de la criatura!", FontTypeNames.FONTTYPE_INFO)
            
                    Exit Sub

                End If
                
290             If .flags.AdminInvisible = 1 Then Exit Sub
            
292             If .Clase = eClass.Druid Then
                    'copio el char original al mimetizado
            
294                 .CharMimetizado.Body = .Char.Body
296                 .CharMimetizado.Head = .Char.Head
298                 .CharMimetizado.CascoAnim = .Char.CascoAnim
300                 .CharMimetizado.ShieldAnim = .Char.ShieldAnim
302                 .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
304                 .flags.Mimetizado = 1
306                 .ShowName = False
308                 .flags.Ignorado = True
                
                    'ahora pongo lo del NPC.
310                 .Char.Body = Npclist(NpcIndex).Char.Body
312                 .Char.Head = Npclist(NpcIndex).Char.Head
314                 .Char.CascoAnim = NingunCasco
316                 .Char.ShieldAnim = NingunEscudo
318                 .Char.WeaponAnim = NingunArma
                      
320                 Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
322                 Call RefreshCharStatus(UserIndex)
                Else
324                 Call WriteConsoleMsg(UserIndex, "Sólo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
    
326             Call InfoHechizo(UserIndex)
328             HechizoCasteado = True
            End With

        End If

        '<EhFooter>
        Exit Sub

HechizoEstadoNPC_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HechizoEstadoNPC " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub HechizoPropNPC(ByVal SpellIndex As Integer, _
                   ByVal NpcIndex As Integer, _
                   ByVal UserIndex As Integer, _
                   ByRef HechizoCasteado As Boolean)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 18/09/2010
        'Handles the Spells that afect the Life NPC
        '14/08/2007 Pablo (ToxicWaste) - Orden general.
        '18/09/2010: ZaMa - Ahora valida si podes ayudar a un npc.
        '***************************************************
        '<EhHeader>
        On Error GoTo HechizoPropNPC_Err

        '</EhHeader>

        Dim daño As Long

100     With Npclist(NpcIndex)

            'Salud
102         If Hechizos(SpellIndex).SubeHP = 1 Then
        
104             HechizoCasteado = CanSupportNpc(UserIndex, NpcIndex)
        
106             If HechizoCasteado Then
        
108                 If .Hostile = 0 Then
110                     Call WriteConsoleMsg(UserIndex, "No puedes curar a la criatura", FontTypeNames.FONTTYPE_INFORED)
112                     HechizoCasteado = False

                        Exit Sub

                    End If
            
114                 daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
116                 daño = daño + Porcentaje(daño, 3 * (UserList(UserIndex).Stats.Elv))
            
118                 Call InfoHechizo(UserIndex)
120                 .Stats.MinHp = .Stats.MinHp + daño

122                 If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
124                 Call WriteConsoleMsg(UserIndex, "Has curado " & daño & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
126                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_CurarSpell))

                End If
        
128         ElseIf Hechizos(SpellIndex).SubeHP = 2 Then

130             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
132                 HechizoCasteado = False

                    Exit Sub

                End If
        
134             Call NPCAtacado(NpcIndex, UserIndex)
136             daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
138             daño = daño + Porcentaje(daño, 3 * (UserList(UserIndex).Stats.Elv))
            
140             If Hechizos(SpellIndex).StaffAffected Then
142                 If UserList(UserIndex).Clase = eClass.Mage Then
144                     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
146                         daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                            'Aumenta daño segun el staff-
                            'Daño = (Daño* (70 + BonifBáculo)) / 100
                        Else
148                         daño = daño * 0.7 'Baja daño a 70% del original

                        End If

                    End If

                End If
        
150             If .NPCtype = DRAGON Then
152                 If UserList(UserIndex).Invent.WeaponEqpObjIndex = VaraMataDragonesIndex Then
154                     daño = daño * 3

                    End If

                End If
        
                'Esta con gran poder?
156             If Power.UserIndex = UserIndex Then
158                 daño = daño * 1.2

                End If
            
160             If UserList(UserIndex).Invent.MagicObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.MagicObjIndex = ANILLOMAGICO Then
162                 daño = daño * 1.04  'laud magico de los bardos 4%

                End If
        
                #If Testeo = 1 Then
                    If EsAdmin(UCase$(UserList(UserIndex).Name)) Then
                        daño = .Stats.MaxHp
                    End If
                #End If
        
164             Call InfoHechizo(UserIndex)
166             HechizoCasteado = True
        
168             If .flags.Snd2 > 0 Then
170                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(.flags.Snd2, .Pos.X, .Pos.Y, .Char.charindex))

                End If
        
                'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
172             daño = daño - .Stats.defM
        
174             If daño < 0 Then daño = 0
                
                Call CalcularDarExp(UserIndex, NpcIndex, daño)
                Call Quests_AddNpc(UserIndex, NpcIndex, daño)
                  
176             .Stats.MinHp = .Stats.MinHp - daño
180             Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoNpcSpell))

182             If .Stats.MinHp < 1 Then
184                 .Stats.MinHp = 0
186                 Call MuereNpc(NpcIndex, UserIndex)

                End If

            End If

        End With

        '<EhFooter>
        Exit Sub

HechizoPropNPC_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoPropNPC " & "at line " & Erl
        
        '</EhFooter>
End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo InfoHechizo_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 25/07/2009
        '25/07/2009: ZaMa - Code improvements.
        '25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
        '***************************************************
        Dim SpellIndex As Integer

        Dim tUser      As Integer

        Dim tNpc       As Integer

100     Dim Valid      As Boolean: Valid = True
    
102     With UserList(UserIndex)
104         SpellIndex = .flags.Hechizo
                
                If Hechizos(SpellIndex).AutoLanzar = 1 Then
                    tUser = UserIndex
                Else
                    tUser = .flags.TargetUser
                End If
                
108         tNpc = .flags.TargetNPC
        
110         If .flags.SlotEvent > 0 Then
112             If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
114                 Valid = False
                End If
            End If
        
116         If Valid Then Call DecirPalabrasMagicas(Hechizos(SpellIndex).PalabrasMagicas, UserIndex)
        
118         If tUser > 0 Then
                ' bueno hace eso para todos como primer paso, avismae cuando lo hayas hecho joya avisme por face cuando lo termines sisi
                ' Los admins invisibles no producen sonidos ni fx's
            
120             If .flags.AdminInvisible = 1 And UserIndex = tUser Then
122                 Call SendData(ToOne, UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
124                 Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y, UserList(tUser).Char.charindex))
                Else
126                 Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
128                 Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y, UserList(tUser).Char.charindex))
                
                End If

130         ElseIf tNpc > 0 Then
132             Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateFX(Npclist(tNpc).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
134             Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y, Npclist(tNpc).Char.charindex))
            
            End If

        End With

        '<EhFooter>
        Exit Sub

InfoHechizo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.InfoHechizo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Hechizos que causan efectos sobre el Area
Public Function HechizoPropAreaUsuario(ByVal UserIndex As Integer) As Boolean

    Dim SpellIndex  As Integer

    Dim Damage      As Long

    Dim TargetIndex As Integer
    
    Dim A           As Long
    
    Dim X           As Byte, Y As Byte
    
    Dim Spell       As tHechizo
    
    With UserList(UserIndex)
        SpellIndex = .flags.Hechizo
        Spell = Hechizos(SpellIndex)
        
        For X = .Pos.X - Spell.AreaX To .Pos.X + Spell.AreaX
            For Y = .Pos.Y - Spell.AreaY To .Pos.Y + Spell.AreaY
                TargetIndex = MapData(.Pos.Map, X, Y).UserIndex
                
                If TargetIndex > 0 Then
                    ' @ Quita SALUD
                    If Spell.SubeHP = 2 And TargetIndex <> UserIndex Then
                        If PuedeAtacar(UserIndex, TargetIndex) Then
                            If HechizoUserReceiveDamage(UserIndex, SpellIndex) Then
                                
                                HechizoPropAreaUsuario = True
                                Damage = HechizoUserUpdateDamage(UserIndex, TargetIndex, SpellIndex)
                                
                                If Damage > 0 Then
                                   ' Call InfoHechizo(UserIndex)
                                    
                                    With UserList(TargetIndex)
                                        .Stats.MinHp = .Stats.MinHp - Damage
                                        Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
                                        Call SubirSkill(TargetIndex, eSkill.Resistencia, True)
                                        Call WriteUpdateHP(TargetIndex)
                                        
                                        Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Damage, d_DañoUserSpell))
                                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
            
                                        'Muere
                                        If .Stats.MinHp < 1 Then
                                            If .flags.AtacablePor <> UserIndex Then Call ContarMuerte(TargetIndex, UserIndex)
                
                                            .Stats.MinHp = 0
                                            Call ActStats(TargetIndex, UserIndex)
                                            Call UserDie(TargetIndex, UserIndex)
    
                                        End If
    
                                    End With
                                End If
                                
                            End If

                        End If

                    End If

                End If

            Next Y
        
        Next X
        

        ' Effects User
         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, X, Y))
         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, UserList(UserIndex).Char.charindex, vbCyan))
         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(.Pos.X, .Pos.Y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
    End With

    
End Function

' @ Comprueba que el usuario pueda recibir un ataque de daño mágico
Public Function HechizoUserReceiveDamage(ByVal UserIndex As Integer, ByVal SpellIndex As Integer) As Boolean
    
    With UserList(UserIndex)
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).config(eConfigEvent.eUseTormenta) = 0 And SpellIndex = eHechizosIndex.eTormenta Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function

            End If
                
            If Events(.flags.SlotEvent).config(eConfigEvent.eUseApocalipsis) = 0 And (SpellIndex = eHechizosIndex.eApocalipsis Or SpellIndex = eHechizosIndex.eExplosionAbismal) Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function

            End If
                
            If Events(.flags.SlotEvent).config(eConfigEvent.eUseDescarga) = 0 And SpellIndex = eHechizosIndex.eDescarga Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function

            End If

        End If
    
    End With

    HechizoUserReceiveDamage = True

End Function

' @ Actualiza el Damage mágico del poder
Public Function HechizoUserUpdateDamage(ByVal UserIndex As Integer, _
                                        ByVal TargetIndex As Integer, _
                                        ByVal SpellIndex As Integer) As Long
    
    Dim Damage As Long
    
    With UserList(TargetIndex)
        Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        
        Damage = Damage + Porcentaje(Damage, 3 * (UserList(UserIndex).Stats.Elv))
        
        If Hechizos(SpellIndex).StaffAffected Then
            If UserList(UserIndex).Clase = eClass.Mage Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Damage = (Damage * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                Else
                    Damage = Damage * 0.7 'Baja Damage a 70% del original

                End If

            End If

        End If
    
        If UserList(UserIndex).Invent.MagicObjIndex = LAUDMAGICO Then
            Damage = Damage * 1.05  'laud magico de los bardos y anillos de druidas

        End If
                
        If UserList(UserIndex).Invent.MagicObjIndex = ANILLOMAGICO Then
            Damage = Damage * 1.03  'laud magico de los bardos y anillos de druidas

        End If

        'cascos antimagia
        If (.Invent.CascoEqpObjIndex > 0) Then
            Damage = Damage - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

        End If
        
        'If .Invent.EscudoEqpObjIndex > 0 Then
        'Damage = Damage - RandomNumber(ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMax)
        'End If
                
        'If .Invent.ArmourEqpObjIndex > 0 Then
        'Damage = Damage - RandomNumber(ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        'End If
                
        'anillos
        If (.Invent.AnilloEqpObjIndex > 0) Then
            Damage = Damage - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

        End If
            
        'Esta con gran poder?
        If Power.UserIndex = UserIndex Then
            Damage = Damage * 1.05

        End If
        
        ' Bonos
        If .flags.SelectedBono > 0 Then
        
            ' Bonos RM
            If ObjData(.flags.SelectedBono).BonoRm > 0 Then
                Damage = Damage * ObjData(.flags.SelectedBono).BonoRm

            End If

        End If
        
        If UserList(UserIndex).flags.SelectedBono > 0 Then
            
            ' Bonos Damage mágicos
            If ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos > 0 Then
                Damage = Damage * ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos

            End If
            
        End If
    
        Damage = Damage - (Damage * .Stats.UserSkills(eSkill.Resistencia) / 2000)
        
        If Damage < 0 Then Damage = 0
        
        HechizoUserUpdateDamage = Damage
    End With

End Function

Public Function HechizoPropUsuario(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 28/04/2010
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
        '***************************************************
        '<EhHeader>
        On Error GoTo HechizoPropUsuario_Err
        '</EhHeader>

        Dim SpellIndex As Integer

        Dim daño As Long

        Dim TargetIndex As Integer

100     SpellIndex = UserList(UserIndex).flags.Hechizo
102     TargetIndex = UserList(UserIndex).flags.TargetUser
      
104     With UserList(TargetIndex)

106         If .flags.Muerto Then
108             Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                Exit Function

            End If
          
            ' <-------- Aumenta Hambre ---------->
110         If Hechizos(SpellIndex).SubeHam = 1 Then
        
112             Call InfoHechizo(UserIndex)
        
114             daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
116             .Stats.MinHam = .Stats.MinHam + daño

118             If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
        
120             If UserIndex <> TargetIndex Then
122                 Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
124                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
126                 Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
        
128             Call WriteUpdateHungerAndThirst(TargetIndex)
    
                ' <-------- Quita Hambre ---------->
130         ElseIf Hechizos(SpellIndex).SubeHam = 2 Then

132             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
134             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
136             If UserIndex <> TargetIndex Then
138                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
                Else

                    Exit Function

                End If
        
140             Call InfoHechizo(UserIndex)
        
142             daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
144             .Stats.MinHam = .Stats.MinHam - daño
        
146             If UserIndex <> TargetIndex Then
148                 Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
150                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
152                 Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
        
154             If .Stats.MinHam < 1 Then
156                 .Stats.MinHam = 0
158                 .flags.Hambre = 1

                End If
        
160             Call WriteUpdateHungerAndThirst(TargetIndex)

            End If
    
            ' <-------- Aumenta Sed ---------->
162         If Hechizos(SpellIndex).SubeSed = 1 Then
        
164             Call InfoHechizo(UserIndex)
        
166             daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
168             .Stats.MinAGU = .Stats.MinAGU + daño

170             If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
        
172             Call WriteUpdateHungerAndThirst(TargetIndex)
             
174             If UserIndex <> TargetIndex Then
176                 Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
178                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
180                 Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
    
                ' <-------- Quita Sed ---------->
182         ElseIf Hechizos(SpellIndex).SubeSed = 2 Then

184             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
        
186             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
188             If UserIndex <> TargetIndex Then
190                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
        
192             Call InfoHechizo(UserIndex)
        
194             daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
196             .Stats.MinAGU = .Stats.MinAGU - daño
        
198             If UserIndex <> TargetIndex Then
200                 Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
202                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
204                 Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
        
206             If .Stats.MinAGU < 1 Then
208                 .Stats.MinAGU = 0
210                 .flags.Sed = 1

                End If
        
212             Call WriteUpdateHungerAndThirst(TargetIndex)
        
            End If
    
            ' <-------- Aumenta Agilidad ---------->
214         If Hechizos(SpellIndex).SubeAgilidad = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
216             If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
218             Call InfoHechizo(UserIndex)
220             daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
        
222             .flags.DuracionEfecto = 1200
224             .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + daño

226             If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
        
228             .flags.TomoPocion = True
230             Call WriteUpdateDexterity(TargetIndex)
    
                ' <-------- Quita Agilidad ---------->
232         ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then

234             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
        
236             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
238             If UserIndex <> TargetIndex Then
240                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
        
242             Call InfoHechizo(UserIndex)
        
244             .flags.TomoPocion = True
246             daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
248             .flags.DuracionEfecto = 700
250             .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - daño

252             If .Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        
254             Call WriteUpdateDexterity(TargetIndex)

            End If
    
            ' <-------- Aumenta Fuerza ---------->
256         If Hechizos(SpellIndex).SubeFuerza = 1 Then
    
                ' Chequea si el status permite ayudar al otro usuario
258             If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
260             Call InfoHechizo(UserIndex)
262             daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
        
264             .flags.DuracionEfecto = 1200
    
266             .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + daño

268             If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
        
270             .flags.TomoPocion = True
272             Call WriteUpdateStrenght(TargetIndex)
    
                ' <-------- Quita Fuerza ---------->
274         ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then

276             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
278             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
280             If UserIndex <> TargetIndex Then
282                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
        
284             Call InfoHechizo(UserIndex)
        
286             .flags.TomoPocion = True
        
288             daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
290             .flags.DuracionEfecto = 700
292             .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - daño

294             If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        
296             Call WriteUpdateStrenght(TargetIndex)

            End If
    
            ' <-------- Cura salud ---------->
298         If Hechizos(SpellIndex).SubeHP = 1 Then
        
                'Verifica que el usuario no este muerto
300             If .flags.Muerto = 1 Then
302                 Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
            
304             If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Or .flags.Desafiando > 0 Then
306                 Call WriteConsoleMsg(UserIndex, "¡No se permite curar desde donde estás!", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
        
                ' Chequea si el status permite ayudar al otro usuario
308             If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
           
310             If .Stats.MinHp = .Stats.MaxHp Then
312                 Call WriteConsoleMsg(UserIndex, "El personaje está sano", FontTypeNames.FONTTYPE_INFORED)

                    Exit Function

                End If
        
314             daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
316             daño = daño + Porcentaje(daño, 3 * (.Stats.Elv))
        
318             Call InfoHechizo(UserIndex)
    
320             .Stats.MinHp = .Stats.MinHp + daño

322             If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        
324             Call WriteUpdateHP(TargetIndex)
        
326             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
328             Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_CurarSpell))
        
                ' <-------- Quita salud (Daña) ---------->
330         ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        
332             If UserIndex = TargetIndex Then
334                 Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Function

                End If
            
                ' Chequeo de Eventos (Anti Spells)
336             If .flags.SlotEvent > 0 Then
338                 If Events(.flags.SlotEvent).config(eConfigEvent.eUseTormenta) = 0 And SpellIndex = eHechizosIndex.eTormenta Then
340                     Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Function

                    End If
                
342                 If Events(.flags.SlotEvent).config(eConfigEvent.eUseApocalipsis) = 0 And (SpellIndex = eHechizosIndex.eApocalipsis Or SpellIndex = eHechizosIndex.eExplosionAbismal) Then
344                     Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Function

                    End If
                
346                 If Events(.flags.SlotEvent).config(eConfigEvent.eUseDescarga) = 0 And SpellIndex = eHechizosIndex.eDescarga Then
348                     Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Function

                    End If

                End If

350             daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        
352             daño = daño + Porcentaje(daño, 3 * (UserList(UserIndex).Stats.Elv))
        
354             If Hechizos(SpellIndex).StaffAffected Then
356                 If UserList(UserIndex).Clase = eClass.Mage Then
358                     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
360                         daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                        Else
362                         daño = daño * 0.7 'Baja daño a 70% del original

                        End If

                    End If

                End If

364             If UserList(UserIndex).Invent.MagicObjIndex = LAUDMAGICO Then
366                 daño = daño * 1.05  'laud magico de los bardos y anillos de druidas

                End If
                
368             If UserList(UserIndex).Invent.MagicObjIndex = ANILLOMAGICO Then
370                 daño = daño * 1.03  'laud magico de los bardos y anillos de druidas

                End If

                'cascos antimagia
372             If (.Invent.CascoEqpObjIndex > 0) Then
374                 daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

                End If
                
                
                If .Pos.Map <> 130 And .Pos.Map <> 131 And .Pos.Map <> 132 Then
                
                    ' Daño mágico para los clanes con CASTILLO NORTE
                    If Castle_CheckBonus(UserList(UserIndex).GuildIndex, eCastle.CASTLE_NORTH) Then
                        daño = daño * 1.02
                    End If
                    
                    ' Resistencia mágica para los clanes con CASTILLO OESTE
                    If Castle_CheckBonus(.GuildIndex, eCastle.CASTLE_WEST) Then
                        daño = daño * 0.98
                    End If
                    
                End If
                
                'If .Invent.EscudoEqpObjIndex > 0 Then
                'Daño = Daño - RandomNumber(ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMax)
                'End If
                
                'If .Invent.ArmourEqpObjIndex > 0 Then
                'Daño = Daño - RandomNumber(ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMax)
                'End If
                
                'anillos
376             If (.Invent.AnilloEqpObjIndex > 0) Then
378                 daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

                End If
            
                'Esta con gran poder?
380             If Power.UserIndex = UserIndex Then
382                 daño = daño * 1.05

                End If
        
                ' Bonos
384             If .flags.SelectedBono > 0 Then
        
                    ' Bonos RM
386                 If ObjData(.flags.SelectedBono).BonoRm > 0 Then
388                     daño = daño * ObjData(.flags.SelectedBono).BonoRm

                    End If

                End If
        
390             If UserList(UserIndex).flags.SelectedBono > 0 Then
            
                    ' Bonos Daño mágicos
392                 If ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos > 0 Then
394                     daño = daño * ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos

                    End If
            
                End If
        
                ' ReliquiaDrag equipped
                'If UserList(UserIndex).Invent.ReliquiaSlot > 0 Then
                'Daño = Effect_UpdatePorc(UserIndex, Daño)
                'End If
        
                ' ReliquiaDrag equipped
                'If .Invent.ReliquiaSlot > 0 Then
                ' Daño = Effect_UpdatePorc(TargetIndex, Daño)
                'End If
        
396             daño = daño - (daño * UserList(TargetIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
        
398             If daño < 0 Then daño = 0
        
400             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
402             If UserIndex <> TargetIndex Then
                      Call checkHechizosEfectividad(UserIndex, TargetIndex)
404                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
            
406             Call InfoHechizo(UserIndex)
        
                If UserList(UserIndex).flags.SlotEvent > 0 Then
                    Events_Add_Damage UserList(UserIndex).flags.SlotEvent, UserList(UserIndex).flags.SlotUserEvent, daño
                End If
        
408             .Stats.MinHp = .Stats.MinHp - daño
        
410             Call SubirSkill(TargetIndex, eSkill.Resistencia, True)
412             Call WriteUpdateHP(TargetIndex)
        
414             Dim Valid As Boolean: Valid = True

416             If .flags.SlotEvent > 0 Then
418                 If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
420                     Valid = False

                    End If

                End If
        
422             If Valid Then
424                 Call SendData(SendTarget.ToOne, TargetIndex, _
                     PrepareMessageConsoleMsg(UserList(UserIndex).Name & " lanzó " & Hechizos(SpellIndex).Nombre & " -" & daño, FontTypeNames.FONTTYPE_FIGHT))
                
426                 Call SendData(SendTarget.ToOne, UserIndex, _
                     PrepareMessageConsoleMsg(Hechizos(SpellIndex).Nombre & " sobre " & UserList(TargetIndex).Name & " -" & daño, FontTypeNames.FONTTYPE_FIGHT))

                End If

428             Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoUserSpell))
430             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
        
                'Muere
432             If .Stats.MinHp < 1 Then
        
434                 If .flags.AtacablePor <> UserIndex Then
                        'Store it!
                        ' Call Statistics.StoreFrag(UserIndex, TargetIndex)
436                     Call ContarMuerte(TargetIndex, UserIndex)

                    End If
            
438                 .Stats.MinHp = 0
440                 Call ActStats(TargetIndex, UserIndex)
442                 Call UserDie(TargetIndex, UserIndex)
444                 Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))

                End If
        
            End If
    
            ' <-------- Aumenta Mana ---------->
446         If Hechizos(SpellIndex).SubeMana = 1 Then
        
448             Call InfoHechizo(UserIndex)
450             .Stats.MinMan = .Stats.MinMan + daño

452             If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan
        
454             Call WriteUpdateMana(TargetIndex)
        
456             If UserIndex <> TargetIndex Then
458                 Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
460                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
462                 Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
    
                ' <-------- Quita Mana ---------->
464         ElseIf Hechizos(SpellIndex).SubeMana = 2 Then

466             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
468             If UserIndex <> TargetIndex Then
470                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
        
472             Call InfoHechizo(UserIndex)
        
474             If UserIndex <> TargetIndex Then
476                 Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
478                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
480                 Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
        
482             .Stats.MinMan = .Stats.MinMan - daño

484             If .Stats.MinMan < 1 Then .Stats.MinMan = 0
        
486             Call WriteUpdateMana(TargetIndex)
        
            End If
    
            ' <-------- Aumenta Stamina ---------->
488         If Hechizos(SpellIndex).SubeSta = 1 Then
490             Call InfoHechizo(UserIndex)
492             .Stats.MinSta = .Stats.MinSta + daño

494             If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
        
496             Call WriteUpdateSta(TargetIndex)
        
498             If UserIndex <> TargetIndex Then
500                 Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
502                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
504                 Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
        
                ' <-------- Quita Stamina ---------->
506         ElseIf Hechizos(SpellIndex).SubeSta = 2 Then

508             If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
510             If UserIndex <> TargetIndex Then
512                 Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
        
514             Call InfoHechizo(UserIndex)
        
516             If UserIndex <> TargetIndex Then
518                 Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
520                 Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Else
522                 Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                End If
        
524             .Stats.MinSta = .Stats.MinSta - daño
        
526             If .Stats.MinSta < 1 Then .Stats.MinSta = 0
        
528             Call WriteUpdateSta(TargetIndex)
        
            End If

        End With

530     HechizoPropUsuario = True

532     Call FlushBuffer(TargetIndex)

        '<EhFooter>
        Exit Function

HechizoPropUsuario_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.HechizoPropUsuario " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, _
                               ByVal TargetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/04/2010
    'Checks if caster can cast support magic on target user.
    '***************************************************
     
    On Error GoTo ErrHandler
 
    With UserList(CasterIndex)
        
        ' Te podes curar a vos mismo
        If CasterIndex = TargetIndex Then
            CanSupportUser = True

            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, TargetIndex) = TRIGGER6_PERMITE Then
            CanSupportUser = True

            Exit Function

        End If
     
        ' Victima criminal?
        If Escriminal(TargetIndex) Then
        
            ' Casteador Ciuda?
            If Not Escriminal(CasterIndex) Then
            
                ' Armadas no pueden ayudar
                If esArmada(CasterIndex) Then
                    Call WriteConsoleMsg(CasterIndex, "Los miembros del ejército real no pueden ayudar a los criminales.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
                
                ' Si el ciuda tiene el seguro puesto no puede ayudar
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                Else

                    ' Penalizacion
                    If DoCriminal Then
                        Call VolverCriminal(CasterIndex)
                    Else
                        Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)
                    End If
                End If
            End If
            
            ' Victima ciuda o army
        Else

            ' Casteador es caos? => No Pueden ayudar ciudas
            If esCaos(CasterIndex) Then
                Call WriteConsoleMsg(CasterIndex, "Los miembros de la legión oscura no pueden ayudar a los ciudadanos.", FontTypeNames.FONTTYPE_INFO)

                Exit Function
                
                ' Casteador ciuda/army?
            ElseIf Not Escriminal(CasterIndex) Then
                
                ' Esta en estado atacable?
                If UserList(TargetIndex).flags.AtacablePor > 0 Then
                    
                    ' No esta atacable por el casteador?
                    If UserList(TargetIndex).flags.AtacablePor <> CasterIndex Then
                    
                        ' Si es armada no puede ayudar
                        If esArmada(CasterIndex) Then
                            Call WriteConsoleMsg(CasterIndex, "Los miembros del ejército real no pueden ayudar a ciudadanos en estado atacable.", FontTypeNames.FONTTYPE_INFO)

                            Exit Function

                        End If
    
                        ' Seguro puesto?
                        If .flags.Seguro Then
                            Call WriteConsoleMsg(CasterIndex, "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", FontTypeNames.FONTTYPE_INFO)

                            Exit Function

                        Else
                            Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)
                        End If
                    End If
                End If
    
            End If
        End If

    End With
    
    CanSupportUser = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", TargetIndex: " & TargetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, _
                       ByVal UserIndex As Integer, _
                       ByVal Slot As Byte)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UpdateUserHechizos_Err
        '</EhHeader>

        Dim LoopC As Byte

100     With UserList(UserIndex)

            'Actualiza un solo slot
102         If Not UpdateAll Then

                'Actualiza el inventario
104             If .Stats.UserHechizos(Slot) > 0 Then
106                 Call ChangeUserHechizo(UserIndex, Slot, .Stats.UserHechizos(Slot))
                Else
108                 Call ChangeUserHechizo(UserIndex, Slot, 0)
                End If

            Else

                'Actualiza todos los slots
110             For LoopC = 1 To MAXUSERHECHIZOS

                    'Actualiza el inventario
112                 If .Stats.UserHechizos(LoopC) > 0 Then
114                     Call ChangeUserHechizo(UserIndex, LoopC, .Stats.UserHechizos(LoopC))
                    Else
116                     Call ChangeUserHechizo(UserIndex, LoopC, 0)
                    End If
            
118             Next LoopC

            End If

        End With

        '<EhFooter>
        Exit Sub

UpdateUserHechizos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.UpdateUserHechizos " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, _
                              ByVal TargetIndex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'Checks if caster can cast support magic on target Npc.
    '***************************************************
     
    On Error GoTo ErrHandler
 
    Dim OwnerIndex As Integer
 
    With UserList(CasterIndex)
        
        OwnerIndex = Npclist(TargetIndex).Owner
        
        ' Si no tiene dueño puede
        If OwnerIndex = 0 Then
            CanSupportNpc = True

            Exit Function

        End If
        
        ' Puede hacerlo si es su propio npc
        If CasterIndex = OwnerIndex Then
            CanSupportNpc = True

            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, OwnerIndex) = TRIGGER6_PERMITE Then
            CanSupportNpc = True

            Exit Function

        End If
     
        ' Victima criminal?
        If Escriminal(OwnerIndex) Then

            ' Victima caos?
            If esCaos(OwnerIndex) Then

                ' Atacante caos?
                If esCaos(CasterIndex) Then
                    ' No podes ayudar a un npc de un caos si sos caos
                    Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
            End If
        
            ' Uno es caos y el otro no, o la victima es pk, entonces puede ayudar al npc
            CanSupportNpc = True

            Exit Function
                
            ' Victima ciuda
        Else

            ' Atacante ciuda?
            If Not Escriminal(CasterIndex) Then

                ' Atacante armada?
                If esArmada(CasterIndex) Then

                    ' Victima armada?
                    If esArmada(OwnerIndex) Then
                        ' No podes ayudar a un npc de un armada si sos armada
                        Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    End If
                End If
                
                ' Uno es armada y el otro ciuda, o los dos ciudas, puede atacar si no tiene seguro
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar a criaturas que luchan contra ciudadanos debes sacarte el seguro.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
                
            End If
            
            ' Atacante criminal y victima ciuda, entonces puede ayudar al npc
            CanSupportNpc = True

            Exit Function
            
        End If
    
    End With
    
    CanSupportNpc = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportNpc, Error: " & Err.number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function

Sub ChangeUserHechizo(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Hechizo As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ChangeUserHechizo_Err
        '</EhHeader>
    
100     UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
    
102     If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
104         Call WriteChangeSpellSlot(UserIndex, Slot)
        Else
106         Call WriteChangeSpellSlot(UserIndex, Slot)
        End If

        '<EhFooter>
        Exit Sub

ChangeUserHechizo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.ChangeUserHechizo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo DisNobAuBan_Err
        '</EhHeader>

        'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
        Dim EraCriminal As Boolean

100     EraCriminal = Escriminal(UserIndex)
    
102     With UserList(UserIndex)

            'Si estamos en la arena no hacemos nada
104         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
106         If .flags.Privilegios And (PlayerType.User) Then
                'pierdo nobleza...
108             .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts

110             If .Reputacion.NobleRep < 0 Then
112                 .Reputacion.NobleRep = 0
                End If
            
                'gano bandido...
114             .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts

116             If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
118             Call WriteMultiMessage(UserIndex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)

120             If Escriminal(UserIndex) Then
122                 If .Faction.Status = r_Armada Then
124                     Call mFacciones.Faction_RemoveUser(UserIndex)
                    Else
126                     Call Guilds_CheckAlineation(UserIndex, a_Neutral)
                    End If
                End If
            End If
        
128         If Not EraCriminal And Escriminal(UserIndex) Then
130             Call RefreshCharStatus(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

DisNobAuBan_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.DisNobAuBan " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub Events_GranBestia_AttackUsers(ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo Events_GranBestia_AttackUsers_Err
        '</EhHeader>

        Dim TempX        As Integer

        Dim TempY        As Integer
    
        Dim X            As Integer

        Dim Y            As Integer
    
        Dim Damage       As Long
    
        Dim UserIndex    As Integer
    
        Dim Attacks      As Byte
    
        Const SpellIndex As Byte = 51

        Const MAX_ATTACK As Byte = 4
    
        Const Map        As Byte = 65

        Const MIN_X      As Byte = 48

        Const MAX_X      As Byte = 62

        Const MIN_Y      As Byte = 43

        Const MAX_Y      As Byte = 54
    
100     X = RandomNumber(MIN_X, MAX_X)
102     Y = RandomNumber(MIN_Y, MAX_Y)
    
104     With Npclist(NpcIndex)

106         For TempX = X To RandomNumber(MAX_X - 3, MAX_X)
108             For TempY = Y To RandomNumber(MAX_Y - 3, MAX_Y)

110                 If InMapBounds(Map, TempX, TempY) Then
112                     UserIndex = MapData(Map, TempX, TempY).UserIndex
                    
114                     If UserIndex > 0 Then

116                         With UserList(UserIndex)

118                             If .flags.SlotEvent > 0 Then
                            
120                                 If RandomNumber(1, 100) <= 10 Then
122                                     Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp) + RandomNumber(50, 100)
                                    Else
124                                     Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                                    End If
                                
126                                 .Stats.MinHp = .Stats.MinHp - Damage
                                
128                                 Call WriteUpdateHP(UserIndex)
130                                 Call WriteConsoleMsg(UserIndex, "¡La gran bestia te ha quitado " & Damage & " puntos de vida!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
132                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
134                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y, .Char.charindex))
                                
136                                 If .Stats.MinHp <= 0 Then
138                                     .Stats.MinHp = .Stats.MaxHp
140                                     Events_GranBestia_MuereUser (UserIndex)

                                        Exit Sub

                                    End If
                            
                                End If

                            End With
                        
142                         Attacks = Attacks + 1
                        End If
                    
144                     If Attacks = MAX_ATTACK Then Exit Sub
                    End If

146             Next TempY
148         Next TempX

        End With

        '<EhFooter>
        Exit Sub

Events_GranBestia_AttackUsers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.Events_GranBestia_AttackUsers " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub checkHechizosEfectividad(ByVal UserIndex As Integer, ByVal TargetUser As Integer)
        '<EhHeader>
        On Error GoTo checkHechizosEfectividad_Err
        '</EhHeader>
100     With UserList(UserIndex)
102         If .Pos.Map = 1 Then Exit Sub

        
104         If UserList(TargetUser).flags.Inmovilizado + UserList(TargetUser).flags.Paralizado = 0 Then
106             .Counters.controlHechizos.HechizosCasteados = .Counters.controlHechizos.HechizosCasteados + 1
        
                Dim efectividad As Double
            
108             efectividad = (100 * .Counters.controlHechizos.HechizosCasteados) / .Counters.controlHechizos.HechizosTotales
            
110             If efectividad >= 85 And .Counters.controlHechizos.HechizosTotales >= 10 Then
112                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El usuario " & .Name & " está lanzando hechizos con una efectividad de " & efectividad & "% (Casteados: " & .Counters.controlHechizos.HechizosCasteados & "/" & .Counters.controlHechizos.HechizosTotales & "), revisar.", FontTypeNames.FONTTYPE_TALK))
114                  Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "El usuario " & .Name & " con IP: " & .IpAddress & " está lanzando hechizos con una efectividad de " & efectividad & "% (Casteados: " & .Counters.controlHechizos.HechizosCasteados & "/" & .Counters.controlHechizos.HechizosTotales & "), revisar.")
                End If
            
           
            Else
116             .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales - 1
            End If
        End With
        '<EhFooter>
        Exit Sub

checkHechizosEfectividad_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modHechizos.checkHechizosEfectividad " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

