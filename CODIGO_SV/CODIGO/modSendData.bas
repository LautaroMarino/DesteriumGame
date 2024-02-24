Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget

    ToAll = 1
    ToOne
    toMap
    toMapSecure
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToConsejoYCaos
    ToClanArea
    ToConsejoCaos
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs
    ToFaction
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, _
                    ByVal sndIndex As Integer, _
                    ByVal sndData As String, _
                    Optional ByVal IsDenounce As Boolean = False, _
                    Optional ByVal IsUrgent As Boolean = False)
        
        
        '<EhHeader>
        On Error GoTo OnError
        '</EhHeader>

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
        'Last Modify Date: 14/11/2010
        'Last modified by: ZaMa
        '14/11/2010: ZaMa - Now denounces can be desactivated.
        '**************************************************************


        Dim LoopC As Long
    
100     Select Case sndRoute

            Case SendTarget.ToOne

102             If UserList(sndIndex).ConnIDValida Then
104                 Call Server.Send(sndIndex, IsUrgent, Writer)
                End If
            
106         Case SendTarget.ToPCArea
108             Call SendToUserArea(sndIndex, sndData)
        
110         Case SendTarget.ToGM

112             For LoopC = 1 To LastUser

114                 If UserList(LoopC).ConnIDValida Then
116                     If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

                            ' Denounces can be desactivated
118                         If IsDenounce Then
120                             If UserList(LoopC).flags.SendDenounces Then
122                                 Call Server.Send(LoopC, False, Writer)
                                End If

                            Else
124                             Call Server.Send(LoopC, False, Writer)
                            End If
                        End If
                    End If

126             Next LoopC
            
128         Case SendTarget.ToAdmins

130             For LoopC = 1 To LastUser

132                 If UserList(LoopC).ConnIDValida Then
134                     If UserList(LoopC).flags.Privilegios And (PlayerType.Admin) Then
136                         If EsGmPriv(LoopC) Then
                                ' Denou(ces can be desactivated
138                             If IsDenounce Then
140                                 If UserList(LoopC).flags.SendDenounces Then
142                                     Call Server.Send(LoopC, False, Writer)
                                    End If
    
                                Else
144                                 Call Server.Send(LoopC, False, Writer)
                                End If
                            End If
                        End If
                    End If

146             Next LoopC
        
148         Case SendTarget.ToAll

150             For LoopC = 1 To LastUser

152                 If UserList(LoopC).ConnIDValida Then
154                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
156                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

158             Next LoopC
        
160         Case SendTarget.ToAllButIndex

162             For LoopC = 1 To LastUser

164                 If (UserList(LoopC).ConnIDValida) And (LoopC <> sndIndex) Then
166                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
168                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

170             Next LoopC
        
172         Case SendTarget.toMap
174             Call SendToMap(sndIndex, sndData)
        
176         Case SendTarget.toMapSecure
178             For LoopC = 1 To LastUser
180                 If (UserList(LoopC).ConnIDValida) And (LoopC <> sndIndex) Then
182                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
184                         If Not MapInfo(UserList(LoopC).Pos.Map).Pk Then
186                             Call Server.Send(LoopC, False, Writer)
                            End If
                        End If
                    End If
188             Next LoopC
                    
190         Case SendTarget.ToGuildMembers
                'LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

                'While LoopC > 0

                'If (UserList(LoopC).ConnIDValida) Then
                'Call Server.send(LoopC, false, Writer)
                'End If

                'LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

                'Wend
        
192         Case SendTarget.ToDeadArea
194             Call SendToDeadUserArea(sndIndex, sndData)
        
196         Case SendTarget.ToPCAreaButIndex
198             Call SendToUserAreaButindex(sndIndex, sndData)
        
200         Case SendTarget.ToClanArea
202             Call SendToUserGuildArea(sndIndex, sndData)
        
204         Case SendTarget.ToPartyArea
206             Call SendToUserPartyArea(sndIndex, sndData)
        
208         Case SendTarget.ToAdminsAreaButConsejeros
210             Call SendToAdminsButConsejerosArea(sndIndex, sndData)
        
212         Case SendTarget.ToNPCArea
214             Call SendToNpcArea(sndIndex, sndData)

216         Case SendTarget.ToDiosesYclan

218             For LoopC = 1 To MAX_GUILD_MEMBER
                
220                 If (UserList(LoopC).ConnIDValida) And (GuildsInfo(sndIndex).Members(LoopC).UserIndex > 0) Then
222                     If UserList(GuildsInfo(sndIndex).Members(LoopC).UserIndex).flags.UserLogged Then 'Esta logeado como usuario?
224                         Call Server.Send(GuildsInfo(sndIndex).Members(LoopC).UserIndex, False, Writer)
                        End If
                    End If

226             Next LoopC

228         Case SendTarget.ToConsejoYCaos

230             For LoopC = 1 To LastUser

232                 If (UserList(LoopC).ConnIDValida) Then
234                     If UserList(LoopC).flags.Privilegios And (PlayerType.RoyalCouncil Or PlayerType.RoyalCouncil) Then
236                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

238             Next LoopC
            
240         Case SendTarget.ToConsejo

242             For LoopC = 1 To LastUser

244                 If (UserList(LoopC).ConnIDValida) Then
246                     If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
248                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

250             Next LoopC
        
252         Case SendTarget.ToConsejoCaos

254             For LoopC = 1 To LastUser

256                 If (UserList(LoopC).ConnIDValida) Then
258                     If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
260                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

262             Next LoopC
        
264         Case SendTarget.ToCiudadanos

266             For LoopC = 1 To LastUser

268                 If (UserList(LoopC).ConnIDValida) Then
270                     If Not Escriminal(LoopC) Then
272                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

274             Next LoopC
        
276         Case SendTarget.ToCriminales

278             For LoopC = 1 To LastUser

280                 If (UserList(LoopC).ConnIDValida) Then
282                     If Escriminal(LoopC) Then
284                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

286             Next LoopC
        
288         Case SendTarget.ToReal

290             For LoopC = 1 To LastUser

292                 If (UserList(LoopC).ConnIDValida) Then
294                     If UserList(LoopC).Faction.Status = r_Armada Then
296                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

298             Next LoopC
        
300         Case SendTarget.ToCaos

302             For LoopC = 1 To LastUser

304                 If (UserList(LoopC).ConnIDValida) Then
306                     If UserList(LoopC).Faction.Status = r_Caos Then
308                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

310             Next LoopC
        
312         Case SendTarget.ToCiudadanosYRMs

314             For LoopC = 1 To LastUser

316                 If (UserList(LoopC).ConnIDValida) Then
318                     If Not Escriminal(LoopC) Then
320                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

322             Next LoopC
        
324         Case SendTarget.ToCriminalesYRMs

326             For LoopC = 1 To LastUser

328                 If (UserList(LoopC).ConnIDValida) Then
330                     If Escriminal(LoopC) Then
332                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

334             Next LoopC
        
336         Case SendTarget.ToRealYRMs

338             For LoopC = 1 To LastUser

340                 If (UserList(LoopC).ConnIDValida) Then
342                     If UserList(LoopC).Faction.Status = r_Armada Then
344                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

346             Next LoopC
        
348         Case SendTarget.ToCaosYRMs

350             For LoopC = 1 To LastUser

352                 If (UserList(LoopC).ConnIDValida) Then
354                     If UserList(LoopC).Faction.Status = r_Caos Then
356                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

358             Next LoopC
        
360         Case SendTarget.ToHigherAdmins

362             For LoopC = 1 To LastUser

364                 If UserList(LoopC).ConnIDValida Then
366                     If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
368                         Call Server.Send(LoopC, False, Writer)
                        End If
                    End If

370             Next LoopC
            
372         Case SendTarget.ToGMsAreaButRmsOrCounselors
374             Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
            
376         Case SendTarget.ToUsersAreaButGMs
378             Call SendToUsersAreaButGMs(sndIndex, sndData)

380         Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
382             Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
        
384         Case SendTarget.ToFaction
386             Call SendToUsersFaction(sndIndex, sndData)
        End Select
    
    
OnError:
        Writer.Clear
        
        If Err.number <> 0 Then
            LogError Err.description & vbCrLf & _
                   "in ServidorArgentum.modSendData.SendData " & _
                   "at line " & Erl
        End If

End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUserArea_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         Call Server.Send(query(i).Name, IsUrgent, Writer)
104     Next i
    
106     Call Server.Send(UserIndex, IsUrgent, Writer)
        '<EhFooter>
        Exit Sub

SendToUserArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUserArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUserAreaButindex_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         Call Server.Send(query(i).Name, IsUrgent, Writer)
104     Next i
        '<EhFooter>
        Exit Sub

SendToUserAreaButindex_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUserAreaButindex " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToDeadUserArea_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         With UserList(query(i).Name)
104             If (.flags.Muerto = 1 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0) Then
106                 Call Server.Send(query(i).Name, IsUrgent, Writer)
                End If
            End With
108     Next i

110     Call Server.Send(UserIndex, IsUrgent, Writer)
        '<EhFooter>
        Exit Sub

SendToDeadUserArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToDeadUserArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUserGuildArea_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
        Dim GuildIndex As Integer
    
100     GuildIndex = UserList(UserIndex).GuildIndex
    
102     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
104         With UserList(query(i).Name)
106             If (.GuildIndex = GuildIndex Or (.flags.Privilegios And PlayerType.Dios)) Then
108                 Call Server.Send(query(i).Name, IsUrgent, Writer)
                End If
            End With
110     Next i

112     Call Server.Send(UserIndex, IsUrgent, Writer)
        '<EhFooter>
        Exit Sub

SendToUserGuildArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUserGuildArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUserPartyArea_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
        Dim GroupIndex As Long
100     GroupIndex = UserList(UserIndex).GroupIndex
102     If GroupIndex = 0 Then Exit Sub
    
104     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
106         If (UserList(query(i).Name).GroupIndex = GroupIndex) Then
108             Call Server.Send(query(i).Name, IsUrgent, Writer)
            End If
110     Next i

112     Call Server.Send(UserIndex, IsUrgent, Writer)
        '<EhFooter>
        Exit Sub

SendToUserPartyArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUserPartyArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToAdminsButConsejerosArea_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         If (UserList(query(i).Name).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
104             Call Server.Send(query(i).Name, IsUrgent, Writer)
            End If
106     Next i
    
108     If (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
110         Call Server.Send(UserIndex, IsUrgent, Writer)
        End If
        '<EhFooter>
        Exit Sub

SendToAdminsButConsejerosArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToAdminsButConsejerosArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToNpcArea_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
102         Call Server.Send(query(i).Name, IsUrgent, Writer)
104     Next i
        '<EhFooter>
        Exit Sub

SendToNpcArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToNpcArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, _
                           ByVal AreaX As Integer, _
                           ByVal AreaY As Integer, _
                           ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToAreaByPos_Err
        '</EhHeader>

        Dim query() As Collision.UUID
        Dim i       As Long
        Dim ItemID  As Long
100     ItemID = Pack(Map, AreaX, AreaY)
    
102     For i = 0 To ModAreas.QueryObservers(ItemID, ENTITY_TYPE_OBJECT, query, ENTITY_TYPE_PLAYER)
104         Call Server.Send(query(i).Name, IsUrgent, Writer)
106     Next i
    
        '<EhFooter>
        Exit Sub

SendToAreaByPos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToAreaByPos " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToMap_Err
        '</EhHeader>
100     Call Server.Broadcast(MapInfo(Map).Players, IsUrgent, Writer)
        '<EhFooter>
        Exit Sub

SendToMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToGMsAreaButRmsOrCounselors_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         With UserList(query(i).Name)
104             If ((.flags.Privilegios And Not PlayerType.User) = .flags.Privilegios) Then
106                 Call Server.Send(query(i).Name, IsUrgent, Writer)
                End If
            End With
108     Next i
    
110     With UserList(UserIndex)
112         If ((.flags.Privilegios And Not PlayerType.User) = .flags.Privilegios) Then
114             Call Server.Send(UserIndex, IsUrgent, Writer)
            End If
        End With
        '<EhFooter>
        Exit Sub

SendToGMsAreaButRmsOrCounselors_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToGMsAreaButRmsOrCounselors " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUsersAreaButGMs_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         If (UserList(query(i).Name).flags.Privilegios And PlayerType.User) Then
104             Call Server.Send(query(i).Name, IsUrgent, Writer)
            End If
106     Next i

108     If (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
110         Call Server.Send(UserIndex, IsUrgent, Writer)
        End If
        '<EhFooter>
        Exit Sub

SendToUsersAreaButGMs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUsersAreaButGMs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUsersAndRmsAndCounselorsAreaButGMs_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         If (UserList(query(i).Name).flags.Privilegios And (PlayerType.User)) Then
104             Call Server.Send(query(i).Name, IsUrgent, Writer)
            End If
106     Next i

108     If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
110         Call Server.Send(UserIndex, IsUrgent, Writer)
        End If
        '<EhFooter>
        Exit Sub

SendToUsersAndRmsAndCounselorsAreaButGMs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUsersAndRmsAndCounselorsAreaButGMs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub SendToUsersFaction(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
        '<EhHeader>
        On Error GoTo SendToUsersAreaButGMs_Err
        '</EhHeader>
        Dim query() As Collision.UUID
        Dim i       As Long
    
100     For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
102         If (UserList(query(i).Name).Faction.Status = UserList(UserIndex).Faction.Status) Then
104             Call Server.Send(query(i).Name, IsUrgent, Writer)
            End If
106     Next i

108     If (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
110         Call Server.Send(UserIndex, IsUrgent, Writer)
        End If
        '<EhFooter>
        Exit Sub

SendToUsersAreaButGMs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSendData.SendToUsersAreaButGMs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo AlertarFaccionarios_Err

        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 17/11/2009
        'Alerta a los faccionarios, dandoles una orientacion
        '**************************************************************
        Dim TempIndex As Integer

        Dim i         As Long

        Dim Font      As FontTypeNames

        Dim query()   As Collision.UUID
        
        With UserList(UserIndex)

100         If esCaos(UserIndex) Then
102             Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
            Else
104             Font = FontTypeNames.FONTTYPE_CONSEJO

            End If
            
            Call SendData(SendTarget.ToFaction, UserIndex, PrepareMessageConsoleMsg("Escuchas el llamado de un líder faccionario que proviene de " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & " " & .Pos.X & " " & .Pos.Y & ")", Font))
        End With

        '<EhFooter>
        Exit Sub

AlertarFaccionarios_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.AlertarFaccionarios " & "at line " & Erl

        

        '</EhFooter>
End Sub
