Attribute VB_Name = "mQuests"
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
'along with this program; if not, you can find it at [url=http://www.affero.org/oagpl.html]http://www.affero.org/oagpl.html[/url]
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at [email=aaron@baronsoft.com]aaron@baronsoft.com[/email]
'for more information about ORE please visit [url=http://www.baronsoft.com/]http://www.baronsoft.com/[/url]
Option Explicit
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 30     'Máxima cantidad de quests que puede tener un usuario al mismo tiempo.
Public NumQuests As Integer
 
Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo FreeQuestSlot_Err
        '</EhHeader>

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Devuelve el próximo slot de quest libre.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
       
100     For i = 1 To MAXUSERQUESTS
            
102         If UserList(UserIndex).QuestStats(i).QuestIndex = 0 Then
104             FreeQuestSlot = i

                Exit Function

            End If

106     Next i
         
108     FreeQuestSlot = 0
        '<EhFooter>
        Exit Function

FreeQuestSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.FreeQuestSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Quest_SetUserPrincipa(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Quest_SetUserPrincipa_Err
        '</EhHeader>
    
        
        Exit Sub
        
100     If UserList(UserIndex).Stats.Elv <= 12 Then
102         Call Quest_SetUser(UserIndex, 1) '  Newbie
        Else
104         Call Quest_SetUser(UserIndex, 2) '  Aledaño
        End If
    
106     Call Quest_SetUser(UserIndex, 10) '  En busca del Oasis :: Misión n°10
108     Call Quest_SetUser(UserIndex, 12) '  La oscuridad bajo la Ciudad :: Misión n°12
110     Call Quest_SetUser(UserIndex, 17) '  Maldición Marabel :: Misión n°17
112     Call Quest_SetUser(UserIndex, 23) '  Tesoro del Dragon Mitico :: Misión n°23
114     Call Quest_SetUser(UserIndex, 26) '  Explorando nuevas Islas :: Misión n°26
116     Call Quest_SetUser(UserIndex, 32) '  Expedición a la Isla Vespar :: Misión n°32
118     Call Quest_SetUser(UserIndex, 37) '  Explorando los Mares Ocultos de Nereo :: Misión n°37
120     Call Quest_SetUser(UserIndex, 40) '  Guerreros del Laberinto Spectra :: Misión n°40
122     Call Quest_SetUser(UserIndex, 44) '  Explorando los Mares de Nueva Esperanza :: Misión n°44
124     Call Quest_SetUser(UserIndex, 51) '  Refugiado en la Isla Veril :: Misión n°51
126     Call Quest_SetUser(UserIndex, 58) '  Isla de los Sacerdotes y Protectores del Rey :: Misión n°58
128     Call Quest_SetUser(UserIndex, 67) '  Descubriendo el Tenebroso Castillo Brezal :: Misión n°67
130     Call Quest_SetUser(UserIndex, 74) '  Afueras del Infierno :: Misión n°74
132     Call Quest_SetUser(UserIndex, 84) '  En busca del Polo Norte :: Misión n°84
        
        
        Call WriteQuestInfo(UserIndex, True, 0)
        Call WriteConsoleMsg(UserIndex, "Misiones> Accede al panel de misiones desde la tecla 'ESC' o bien escribiendo /MISIONES", FontTypeNames.FONTTYPE_CRITICO)
        
        '<EhFooter>
        Exit Sub

Quest_SetUserPrincipa_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.Quest_SetUserPrincipa " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


' Setea una nueva mision/objetivo en el personaje
Public Sub Quest_SetUser(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)

        '<EhHeader>
        On Error GoTo Quest_SetUser_Err

        '</EhHeader>
        
        Dim QuestSlot As Integer
        
        Exit Sub
        QuestSlot = FreeQuestSlot(UserIndex)
        
        If QuestSlot > 0 Then

            With UserList(UserIndex).QuestStats(QuestSlot)
102             .QuestIndex = QuestIndex
        
104             If QuestList(QuestIndex).RequiredNPCs > 0 Then ReDim .NPCsKilled(1 To QuestList(QuestIndex).RequiredNPCs) As Long
106             If QuestList(QuestIndex).RequiredChestOBJs > 0 Then ReDim .ObjsPick(1 To QuestList(QuestIndex).RequiredChestOBJs) As Long
108             If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then ReDim .ObjsSale(1 To QuestList(QuestIndex).RequiredSaleOBJs) As Long
            End With
        
        Else
            Call WriteConsoleMsg(UserIndex, "Error al otorgar una nueva misión.", FontTypeNames.FONTTYPE_INFORED)
        End If
        
100
        '<EhFooter>
        Exit Sub

Quest_SetUser_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quest_SetUser " & "at line " & Erl
        
        '</EhFooter>
End Sub

' Comprueba si tiene los objetos
Public Sub Quests_Check_Objs(ByVal UserIndex As Integer, _
                             ByVal ObjIndex As Integer, _
                             ByVal Amount As Integer)

        '<EhHeader>
        On Error GoTo Quests_Check_ChestObj_Err

        '</EhHeader>

        Dim A          As Long, B As Long

        Dim QuestIndex As Integer
        
        For B = 1 To MAXUSERQUESTS
            QuestIndex = UserList(UserIndex).QuestStats(B).QuestIndex
        
            If QuestIndex = 0 Then Exit Sub
        
100         With QuestList(QuestIndex)

102             If .RequiredOBJs > 0 Then

104                 For A = 1 To .RequiredOBJs

                        If ObjIndex = .RequiredObj(A).ObjIndex Then
                            If TieneObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex) Then
                                Call Quests_Final(UserIndex, B)
                                Exit For
                            Else
                                Call WriteQuestInfo(UserIndex, False, B)
                                
                            End If

                        End If

112                 Next A

                End If
                
            End With
        
        Next B

        '<EhFooter>
        Exit Sub

Quests_Check_ChestObj_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_Check_ChestObj " & "at line " & Erl

        '</EhFooter>
End Sub

' Le otorga la recompensa que merece por haber completado la misión
Public Function Quests_CheckFinish(ByVal UserIndex As Integer, ByVal Slot As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Quests_CheckFinish_Err
        '</EhHeader>

        Dim QuestIndex As Integer
    
100     With UserList(UserIndex)
102         QuestIndex = .QuestStats(Slot).QuestIndex
        
            Dim A As Long
        
104         With QuestList(QuestIndex)

106             If .RequiredNPCs > 0 Then

108                 For A = 1 To .RequiredNPCs

110                     If .RequiredNpc(A).Amount * .RequiredNpc(A).Hp <> UserList(UserIndex).QuestStats(Slot).NPCsKilled(A) Then
                            Exit Function

                        End If

112                 Next A

                End If
            
114              If .RequiredSaleOBJs > 0 Then

116                 For A = 1 To .RequiredSaleOBJs

118                     If .RequiredSaleObj(A).Amount <> UserList(UserIndex).QuestStats(Slot).ObjsSale(A) Then
                            Exit Function

                        End If

120                 Next A

                End If
            
122             If .RequiredChestOBJs > 0 Then

124                 For A = 1 To .RequiredChestOBJs

126                     If .RequiredChestObj(A).Amount <> UserList(UserIndex).QuestStats(Slot).ObjsPick(A) Then
                            Exit Function

                        End If

128                 Next A

                End If
            
130             If .RequiredOBJs > 0 Then

132                 For A = 1 To .RequiredOBJs

134                     If Not TieneObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex) Then
                            Exit Function

                        End If

136                 Next A

                End If
            
            End With

        End With
    
138     Quests_CheckFinish = True

        '<EhFooter>
        Exit Function

Quests_CheckFinish_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.Quests_CheckFinish " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Quests_Reward(ByVal UserIndex As Integer, ByVal Slot As Integer)

        '<EhHeader>
        On Error GoTo Quests_Reward_Err

        '</EhHeader>
    
        Dim A          As Long

        Dim QuestIndex As Integer
    
        Dim Text       As String
        
        'Dim List()     As String
        
        Dim Obj        As Obj
    
    '    ReDim Preserve List(0) As String
        
100     QuestIndex = UserList(UserIndex).QuestStats(Slot).QuestIndex
    
102     With UserList(UserIndex)

104         If QuestList(QuestIndex).RequiredOBJs > 0 Then

                If QuestList(QuestIndex).Remove > 0 Then

106                 For A = 1 To QuestList(QuestIndex).RequiredOBJs
108                     Call QuitarObjetos(QuestList(QuestIndex).RequiredObj(A).ObjIndex, QuestList(QuestIndex).RequiredObj(A).Amount, UserIndex)
110                 Next A

                End If
                
            End If
            
112         If QuestList(QuestIndex).RewardEXP > 0 Then
114             .Stats.Exp = .Stats.Exp + QuestList(QuestIndex).RewardEXP
116             Call CheckUserLevel(UserIndex)
118             Call WriteUpdateExp(UserIndex)
                  
                'ReDim Preserve List(0 To UBound(List) + 1) As String
                'List(1) = "+" & QuestList(QuestIndex).RewardEXP & " EXP"

            End If
            
122         If QuestList(QuestIndex).RewardEldhir > 0 Then
124             .Account.Eldhir = .Account.Eldhir + QuestList(QuestIndex).RewardEldhir
126             Call WriteUpdateDsp(UserIndex)
                  
               ' ReDim Preserve List(0 To UBound(List) + 1) As String
                'List(UBound(List)) = "+" & QuestList(QuestIndex).RewardEldhir & " DSP"

            End If
            
130         If QuestList(QuestIndex).RewardGLD > 0 Then
132             .Stats.Gld = .Stats.Gld + QuestList(QuestIndex).RewardGLD
134             Call WriteUpdateGold(UserIndex)
            
              '  ReDim Preserve List(0 To UBound(List) + 1) As String
               ' List(UBuond(List)) = "+" & QuestList(QuestIndex).RewardGLD & " ORO"

            End If
        
138         If QuestList(QuestIndex).RewardOBJs > 0 Then

142             For A = 1 To QuestList(QuestIndex).RewardOBJs
                    
144                 Obj.ObjIndex = QuestList(QuestIndex).RewardObj(A).ObjIndex
146                 Obj.Amount = QuestList(QuestIndex).RewardObj(A).Amount
                      
                    If ObjData(Obj.ObjIndex).OBJType = otRangeQuest Then
                        Call UseCofrePoder(UserIndex, ObjData(Obj.ObjIndex).Range)
                    Else

                        If ClasePuedeUsarItem(UserIndex, Obj.ObjIndex) Then
                        
                            If Not MeterItemEnInventario(UserIndex, Obj) Then
150                             Call TirarItemAlPiso(.Pos, Obj)
    
                            End If

                        End If

                    End If
                      
                  '  ReDim Preserve List(0 To UBound(List) + 1) As String
                   ' List(UBound(List)) = "+" & ObjData(Obj.ObjIndex).Name & " (x" & QuestList(QuestIndex).RewardObj(A).Amount & ")"
154             Next A
            
            End If
        
            Call WriteUpdateFinishQuest(UserIndex, QuestIndex)
              
158         Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(RandomNumber(eSound.sVictory3, eSound.sVictory5), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
160

        End With
    
        '<EhFooter>
        Exit Sub

Quests_Reward_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_Reward " & "at line " & Erl

        

        '</EhFooter>
End Sub

' Comprueba cuando golpea una criatura
Public Sub Quests_AddNpc(ByVal UserIndex As Integer, _
                         ByVal NpcIndex As Integer, _
                         ByVal Damage As Long)

        '<EhHeader>
        On Error GoTo Quests_AddNpc_Err

        '</EhHeader>
    
        Dim Diferencia As Long

        Dim A          As Long
        
        Dim B As Long
        
        Dim TempQuest As tUserQuest
        
        For B = 1 To MAXUSERQUESTS

100         With UserList(UserIndex).QuestStats(B)
                
102             If .QuestIndex Then
                    
                    TempQuest = UserList(UserIndex).QuestStats(B)
                    
104                 If Damage > Npclist(NpcIndex).Stats.MinHp Then
106                     Diferencia = Abs(Npclist(NpcIndex).Stats.MinHp)
                    Else
108                     Diferencia = Damage

                    End If
        
110                 If QuestList(.QuestIndex).RequiredNPCs Then
        
112                     For A = 1 To QuestList(.QuestIndex).RequiredNPCs
        
114                         If QuestList(.QuestIndex).RequiredNpc(A).NpcIndex = Npclist(NpcIndex).numero Then
116                             .NPCsKilled(A) = .NPCsKilled(A) + Abs(Diferencia)

118                             If .NPCsKilled(A) >= QuestList(.QuestIndex).RequiredNpc(A).Amount * Npclist(NpcIndex).Stats.MaxHp Then .NPCsKilled(A) = QuestList(.QuestIndex).RequiredNpc(A).Amount * Npclist(NpcIndex).Stats.MaxHp
                              
                                Call Quests_Final(UserIndex, B)
                               
                                'Exit For
                        
                            End If

122                     Next A

                    End If

                End If

            End With
        
        Next B

        '<EhFooter>
        Exit Sub

Quests_AddNpc_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_AddNpc " & "at line " & Erl & " in quest: " & TempQuest.QuestIndex

        '</EhFooter>
End Sub

' Comprueba cuando vende un objeto
Public Sub Quests_AddSale(ByVal UserIndex As Integer, _
                          ByVal ObjIndex As Integer, _
                          ByVal Amount As Long)

        '<EhHeader>
        On Error GoTo Quests_AddSale_Err

        '</EhHeader>
    
        Dim Diferencia As Long

        Dim A          As Long
        
        Dim B As Long
        
        For B = 1 To MAXUSERQUESTS

100         With UserList(UserIndex).QuestStats(B)

102             If .QuestIndex Then
104                 If QuestList(.QuestIndex).RequiredSaleOBJs Then
        
106                     For A = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
        
108                         If QuestList(.QuestIndex).RequiredSaleObj(A).ObjIndex = ObjIndex Then
110                             .ObjsSale(A) = .ObjsSale(A) + Amount

112                             If .ObjsSale(A) >= QuestList(.QuestIndex).RequiredSaleObj(A).Amount Then .ObjsSale(A) = QuestList(.QuestIndex).RequiredSaleObj(A).Amount
                              
114                             Call Quests_Final(UserIndex, B)
                                Exit For

                            End If

116                     Next A

                    End If

                End If

            End With

        Next B

        '<EhFooter>
        Exit Sub

Quests_AddSale_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_AddSale " & "at line " & Erl
        
        '</EhFooter>
End Sub

' Comprueba cuando abre un cofre especifico
Public Sub Quests_AddChest(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           ByVal Amount As Long)

        '<EhHeader>
        On Error GoTo Quests_AddChest_Err

        '</EhHeader>
    
        Dim Diferencia As Long

        Dim A          As Long
        
        Dim B          As Long
        
        For B = 1 To MAXUSERQUESTS
         
100         With UserList(UserIndex).QuestStats(B)

102             If .QuestIndex Then
104                 If QuestList(.QuestIndex).RequiredChestOBJs Then
        
106                     For A = 1 To QuestList(.QuestIndex).RequiredChestOBJs
        
108                         If QuestList(.QuestIndex).RequiredChestObj(A).ObjIndex = ObjIndex Then
110                             .ObjsPick(A) = .ObjsPick(A) + Amount

112                             If .ObjsPick(A) >= QuestList(.QuestIndex).RequiredChestObj(A).Amount Then .ObjsPick(A) = QuestList(.QuestIndex).RequiredChestObj(A).Amount
                              
114                             Call Quests_Final(UserIndex, B)
                                Exit For

                            End If

116                     Next A

                    End If

                End If

            End With
        
        Next B
        
        
        '<EhFooter>
        Exit Sub

Quests_AddChest_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_AddChest " & "at line " & Erl
        
        '</EhFooter>
End Sub

' Chequea si pasa la misión
Public Sub Quests_Final(ByVal UserIndex As Integer, ByVal Slot As Integer)
        '<EhHeader>
        On Error GoTo Quests_Final_Err
        '</EhHeader>
100     If UserList(UserIndex).QuestStats(Slot).QuestIndex = 0 Then Exit Sub

102     If Quests_CheckFinish(UserIndex, Slot) Then
104         Call mQuests.Quests_Next(UserIndex, Slot)
        End If
        
         Call WriteQuestInfo(UserIndex, False, Slot)
106
        '<EhFooter>
        Exit Sub

Quests_Final_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.Quests_Final " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
' Tipea la próxima misión que debera cumplir de manera automatica
Public Sub Quests_Next(ByVal UserIndex As Integer, _
                                   ByVal Slot As Integer)
        '<EhHeader>
        On Error GoTo Quests_Next_Err
        '</EhHeader>
    
        Dim NextQuest As Integer
        Dim NextQuest_1 As Integer
    
100     With UserList(UserIndex)
102         NextQuest = QuestList(.QuestStats(Slot).QuestIndex).NextQuest
        
104         Call Quests_Reward(UserIndex, Slot)
              Call CleanQuestSlot(UserList(UserIndex), Slot)
              
106         If NextQuest > 0 Then
108             If QuestList(NextQuest).RequiredNPCs > 0 Then
110                 ReDim .QuestStats(Slot).NPCsKilled(1 To QuestList(NextQuest).RequiredNPCs) As Long
                End If
            
112             If QuestList(NextQuest).RequiredSaleOBJs > 0 Then
114                 ReDim .QuestStats(Slot).ObjsSale(1 To QuestList(NextQuest).RequiredSaleOBJs) As Long
                End If
            
116             If QuestList(NextQuest).RequiredChestOBJs > 0 Then
118                 ReDim .QuestStats(Slot).ObjsPick(1 To QuestList(NextQuest).RequiredChestOBJs) As Long
                End If
            
120              .QuestStats(Slot).QuestIndex = NextQuest
            
122             'NextQuest_1 = QuestList(.QuestStats(Slot).QuestIndex).NextQuest
                
                  'If Len(QuestList(NextQuest).Desc) Then
124                 'Call WriteConsoleMsg(UserIndex, QuestList(NextQuest).Desc, FontTypeNames.FONTTYPE_INFOGREEN)
                  'End If
                  
                  Call WriteQuestInfo(UserIndex, False, Slot)
                  
            End If
    
        End With

        '<EhFooter>
        Exit Sub

Quests_Next_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.Quests_Next " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
 
Public Sub CleanQuestSlot(ByRef IUser As User, ByVal Slot As Integer)

        '<EhHeader>
        On Error GoTo CleanQuestSlot_Err

        '</EhHeader>

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Limpia un slot de quest de un usuario.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer

102     With IUser.QuestStats(Slot)

104         If .QuestIndex Then
106             If QuestList(.QuestIndex).RequiredNPCs Then

108                 For i = 1 To QuestList(.QuestIndex).RequiredNPCs
110                     .NPCsKilled(i) = 0
112                 Next i

                End If

114             If QuestList(.QuestIndex).RequiredChestOBJs Then

116                 For i = 1 To QuestList(.QuestIndex).RequiredChestOBJs
118                     .ObjsPick(i) = 0
120                 Next i

                End If

122             If QuestList(.QuestIndex).RequiredSaleOBJs Then

124                 For i = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
126                     .ObjsSale(i) = 0
128                 Next i

                End If

            End If

130         .QuestIndex = 0

        End With

        '<EhFooter>
        Exit Sub

CleanQuestSlot_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.CleanQuestSlot " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub
 
Public Sub LoadQuests()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga el archivo Quests_FilePath en el array QuestList.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo ErrorHandler

    Dim Reader    As clsIniManager

    Dim tmpStr    As String

    Dim i         As Integer

    Dim j         As Integer
         
    'Cargamos el clsIniManager en memoria
    Set Reader = New clsIniManager
         
    'Lo inicializamos para el archivo Quests_FilePath
    Call Reader.Initialize(Quests_FilePath)
         
    'Redimensionamos el array
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests) As tQuest
         
    'Cargamos los datos
    For i = 1 To NumQuests

        With QuestList(i)
            .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .Desc = Reader.GetValue("QUEST" & i, "Desc")
            .DescFinish = Reader.GetValue("QUEST" & i, "DescFinal")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            .DoneQuestMessage = val(Reader.GetValue("QUEST" & i, "DoneQuestMessage"))
            .RequiredBronce = val(Reader.GetValue("QUEST" & i, "RequiredBronce"))
            .DoneQuest = val(Reader.GetValue("QUEST" & i, "DoneQuest"))
            .RequiredPlata = val(Reader.GetValue("QUEST" & i, "RequiredPlata"))
            .RequiredOro = val(Reader.GetValue("QUEST" & i, "RequiredOro"))
            .RequiredPremium = val(Reader.GetValue("QUEST" & i, "RequiredPremium"))
            .LastQuest = val(Reader.GetValue("QUEST" & i, "LastQuest"))
            .NextQuest = val(Reader.GetValue("QUEST" & i, "NextQuest"))
            
            .Remove = val(Reader.GetValue("QUEST" & i, "Remove"))
            
            .RewardDaily = val(Reader.GetValue("QUEST" & i, "RewardDaily"))
            
            If .RewardDaily > 0 Then
                DailyLast = DailyLast + 1
                
                ReDim Preserve QuestDaily(DailyLast) As Byte
            
                QuestDaily(DailyLast) = i
            End If

            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

            If .RequiredOBJs > 0 Then
                ReDim .RequiredObj(1 To .RequiredOBJs)

                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                         
                    .RequiredObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            
            
            ' Venta de Objetos
              .RequiredSaleOBJs = val(Reader.GetValue("QUEST" & i, "RequiredSaleOBJs"))

            If .RequiredSaleOBJs > 0 Then
                ReDim .RequiredSaleObj(1 To .RequiredSaleOBJs)

                For j = 1 To .RequiredSaleOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredSaleOBJ" & j)
                         
                    .RequiredSaleObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredSaleObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
                
            ' Requiere:: Abrir Cofres de los Mapas
            .RequiredChestOBJs = val(Reader.GetValue("QUEST" & i, "RequiredChestOBJs"))

            If .RequiredChestOBJs > 0 Then
                ReDim .RequiredChestObj(1 To .RequiredChestOBJs)

                For j = 1 To .RequiredChestOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredChestOBJ" & j)
                         
                    .RequiredChestObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredChestObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

            If .RequiredNPCs > 0 Then
                ReDim .RequiredNpc(1 To .RequiredNPCs)

                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                         
                    .RequiredNpc(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNpc(j).Amount = val(ReadField(2, tmpStr, 45))
                    .RequiredNpc(j).Hp = val(LeerNPCs.GetValue("NPC" & .RequiredNpc(j).NpcIndex, "MAXHP"))
                Next j

            End If
                 
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEldhir = val(Reader.GetValue("QUEST" & i, "RewardEldhir"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            
           ' Call WriteVar(Quests_FilePath, "QUEST" & i, "RewardGLD", CStr(.RewardGLD))
          '  Call WriteVar(Quests_FilePath, "QUEST" & i, "RewardEXP", CStr(.RewardEXP))
            
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))

            If .RewardOBJs > 0 Then
                ReDim .RewardObj(1 To .RewardOBJs)

                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                         
                    .RewardObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If

        End With

    Next i
         
   ' Reader.DumpFile Quests_FilePath
    
    'Eliminamos la clase
    Set Reader = Nothing
    
    Call DataServer_Generate_Quests
    
    Exit Sub
                         
ErrorHandler:
    LogError "Error cargando el archivo " & Quests_FilePath
End Sub
 
Public Sub LoadQuestStats(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)

        '<EhHeader>
        On Error GoTo LoadQuestStats_Err

        '</EhHeader>

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Carga las QuestStats del usuario.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        Dim j      As Integer

        Dim tmpStr As String
        
        Dim A      As Long
        
        ' Adaptation Chars
        If val(Userfile.GetValue("QUESTS", "Q1")) = 999 Then
        
            For A = 1 To MAXUSERQUESTS
                Call mQuests.CleanQuestSlot(UserList(UserIndex), A)
            Next A
            
            Call Quest_SetUserPrincipa(UserIndex)
            Exit Sub

        End If
                      
        For A = 1 To MAXUSERQUESTS
                
100         With UserList(UserIndex).QuestStats(A)

102             tmpStr = Userfile.GetValue("QUESTS", "Q" & A)
                 
104             .QuestIndex = val(ReadField(1, tmpStr, 45))
                      
106             If .QuestIndex Then
108                 If QuestList(.QuestIndex).RequiredNPCs Then
            
110                     ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
                         
112                     For j = 1 To QuestList(.QuestIndex).RequiredNPCs
114                         .NPCsKilled(j) = val(ReadField(j + 1, tmpStr, 45))
116                     Next j

                    End If
                
118                 If QuestList(.QuestIndex).RequiredChestOBJs Then
120                     ReDim .ObjsPick(1 To QuestList(.QuestIndex).RequiredChestOBJs)
                         
122                     For j = 1 To QuestList(.QuestIndex).RequiredChestOBJs
124                         .ObjsPick(j) = val(ReadField(QuestList(.QuestIndex).RequiredNPCs + j + 1, tmpStr, 45))
126                     Next j

                    End If
                
128                 If QuestList(.QuestIndex).RequiredSaleOBJs Then
130                     ReDim .ObjsSale(1 To QuestList(.QuestIndex).RequiredSaleOBJs)
                         
132                     For j = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
134                         .ObjsSale(j) = val(ReadField(QuestList(.QuestIndex).RequiredChestOBJs + j + 1, tmpStr, 45))
136                     Next j

                    End If
                    
                End If

            End With
                         
        Next A

        '<EhFooter>
        Exit Sub

LoadQuestStats_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.LoadQuestStats " & "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SaveQuestStats(ByRef IQuest() As tUserQuest, ByRef Manager As clsIniManager)
        '<EhHeader>
        On Error GoTo SaveQuestStats_Err
        '</EhHeader>


        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Guarda las QuestStats del usuario.
        'Last modified: 29/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i            As Integer

        Dim j            As Integer

        Dim tmpStr       As String
       
        Dim TempRequired As String
        
        Dim A            As Long
        
100     For A = 1 To MAXUSERQUESTS

102         With IQuest(A)
104             tmpStr = .QuestIndex
106             TempRequired = vbNullString
                  
108             If .QuestIndex Then
110                 If QuestList(.QuestIndex).RequiredNPCs Then

112                     For j = 1 To QuestList(.QuestIndex).RequiredNPCs
114                         TempRequired = TempRequired & "-" & CStr(.NPCsKilled(j))
116                     Next j
                        
118                     tmpStr = tmpStr & TempRequired
120                     TempRequired = vbNullString

                    End If
                    
122                 If QuestList(.QuestIndex).RequiredChestOBJs Then

124                     For j = 1 To QuestList(.QuestIndex).RequiredChestOBJs
126                         TempRequired = TempRequired & "-" & CStr(.ObjsPick(j))
128                     Next j

130                     tmpStr = tmpStr & TempRequired
132                     TempRequired = vbNullString

                    End If
                    
134                 If QuestList(.QuestIndex).RequiredSaleOBJs Then
                    
136                     For j = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
138                         TempRequired = TempRequired & "-" & CStr(.ObjsSale(j))
140                     Next j
                   
142                     tmpStr = tmpStr & TempRequired
144                     TempRequired = vbNullString

                    End If

                End If
             
146             Call Manager.ChangeValue("QUESTS", "Q" & A, tmpStr)

            End With
        
148     Next A

        '<EhFooter>
        Exit Sub

SaveQuestStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.SaveQuestStats " & _
               "at line " & Erl & " en QuestIndex: " & A
        Resume Next
        '</EhFooter>
End Sub
 
Private Function Quests_SearchQuest(ByVal UserIndex As Integer, _
                                    ByVal QuestIndex As Byte) As Boolean
        '<EhHeader>
        On Error GoTo Quests_SearchQuest_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAXUSERQUESTS

102         With UserList(UserIndex).QuestStats(A)

104             If .QuestIndex = QuestIndex Then
106                 Quests_SearchQuest = True

                    Exit Function

                End If

            End With

108     Next A

        '<EhFooter>
        Exit Function

Quests_SearchQuest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mQuests.Quests_SearchQuest " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
