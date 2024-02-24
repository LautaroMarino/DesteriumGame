Attribute VB_Name = "mEffect"
Option Explicit

Public Sub Effect_Add(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer)
    
    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        If .flags.SlotEvent > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar Scrolls en eventos automáticos!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If ObjData(ObjIndex).Time = 0 Then
            Call Effect_Selected(UserIndex, ObjData(ObjIndex).BonusTipe, ObjData(ObjIndex).BonusValue, Slot)
        Else

            If .Counters.TimeBonus > 0 Then
                Call WriteConsoleMsg(UserIndex, "Ya tienes un efecto activo.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub
'
            End If
            
            .Stats.BonusTipe = ObjData(ObjIndex).BonusTipe
            .Stats.BonusValue = ObjData(ObjIndex).BonusValue
            .Counters.TimeBonus = ObjData(ObjIndex).Time
            
            Call WriteConsoleMsg(UserIndex, "Tendrás el efecto elegido durante " & Int(.Counters.TimeBonus / 60) & " minutos.", FontTypeNames.FONTTYPE_INFOGREEN)
            
            'Quitamos del inv el item
            If ObjData(ObjIndex).RemoveObj > 0 Then
                Call QuitarUserInvItem(UserIndex, Slot, ObjData(ObjIndex).RemoveObj)
                Call UpdateUserInv(False, UserIndex, Slot)
            End If
        End If
        

        
    End With

    Exit Sub

ErrHandler:
End Sub

Private Sub Effect_Selected(ByVal UserIndex As Integer, _
                            ByVal Tipe As eEffectObj, _
                            ByVal Value As Single, _
                            ByVal Slot As Byte)

        '<EhHeader>
        On Error GoTo Effect_Selected_Err

        '</EhHeader>
                            
100     With UserList(UserIndex)

102         Select Case Tipe

                Case eEffectObj.e_Exp
104                 .Stats.Exp = .Stats.Exp + Value
106                 Call CheckUserLevel(UserIndex)
108                 Call WriteUpdateExp(UserIndex)
110                 Call WriteConsoleMsg(UserIndex, "¡Has ganado " & Value & " puntos de experiencia!", FontTypeNames.FONTTYPE_INFOGREEN)
                
112                 Call QuitarUserInvItem(UserIndex, Slot, 1)
114                 Call UpdateUserInv(False, UserIndex, Slot)
            
116             Case eEffectObj.e_Gld
                      
118                 .Stats.Gld = .Stats.Gld + Value

                    If (.Stats.Gld) > MAXORO Then .Stats.Gld = MAXORO
120                 Call WriteUpdateGold(UserIndex)
122                 Call WriteConsoleMsg(UserIndex, "¡Has ganado " & Value & " Monedas de Oro!", FontTypeNames.FONTTYPE_INFOGREEN)
                
124                 Call QuitarUserInvItem(UserIndex, Slot, 1)
126                 Call UpdateUserInv(False, UserIndex, Slot)
                
128             Case eEffectObj.e_Revive
130                 Call RevivirUsuario(UserIndex)
132                 .Stats.MinHam = 0
134                 .Stats.MinAGU = 0
136                 .flags.Hambre = 1
138                 .flags.Sed = 1
140                 Call WriteUpdateHungerAndThirst(UserIndex)
142                 Call WriteConsoleMsg(UserIndex, "¡Has vuelvo al mundo! ¡Quedas sediento!", FontTypeNames.FONTTYPE_INFOGREEN)
                
144                 Call QuitarUserInvItem(UserIndex, Slot, 1)
146                 Call UpdateUserInv(False, UserIndex, Slot)
                
148             Case eEffectObj.e_NewHead, eEffectObj.e_NewHeadClassic

                    Dim TempHead As Integer

150                 TempHead = .Char.Head
                      
152                 Call User_GenerateNewHead(UserIndex, Tipe)
156
                
158                 If TempHead <> .Char.Head Then
160                     Call QuitarUserInvItem(UserIndex, Slot, 1)
162                     Call UpdateUserInv(False, UserIndex, Slot)

                    End If
                    
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                    Call SaveUser(UserList(UserIndex), CharPath & UCase$(.Name) & ".chr", False)
                    
                    Case eEffectObj.e_ChangeGenero
                        
                        If .Genero = Hombre Then
                            .Genero = Mujer
                        Else
                            .Genero = Hombre
                        End If
                        
                                                
                        If .Invent.ArmourEqpObjIndex > 0 Then
                            If Not SexoPuedeUsarItem(UserIndex, .Invent.ArmourEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                            End If
                        Else
                            Call DarCuerpoDesnudo(UserIndex)
                        End If
                        
                        Call User_GenerateNewHead(UserIndex, eEffectObj.e_NewHead)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                        Call SaveUser(UserList(UserIndex), CharPath & UCase$(.Name) & ".chr", False)
                        Call WriteConsoleMsg(UserIndex, "¡Has cambiado tu género!", FontTypeNames.FONTTYPE_INFOGREEN)

                        
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Call UpdateUserInv(False, UserIndex, Slot)
            End Select
    
        End With

        '<EhFooter>
        Exit Sub

Effect_Selected_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mEffect.Effect_Selected " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub Effect_Remove(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Effect_Remove_Err
        '</EhHeader>

100     With UserList(UserIndex)
102         .Stats.BonusTipe = 0
104         .Stats.BonusValue = 0
106         .Counters.TimeBonus = 0
        End With
    
        '<EhFooter>
        Exit Sub

Effect_Remove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mEffect.Effect_Remove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Effect_Loop(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Effect_Loop_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            ' // NUEVO
102         If .Pos.Map = 0 Then Exit Sub
104         If MapInfo(.Pos.Map).Pk = False Then Exit Sub
        
106         If .Counters.TimeBonus > 0 Then
        
108             .Counters.TimeBonus = .Counters.TimeBonus - 1

110             If .Counters.TimeBonus = 0 Then
112                 Effect_Remove (UserIndex)
114                 Call WriteConsoleMsg(UserIndex, "El efecto se ha ido.", FontTypeNames.FONTTYPE_INFORED)
                End If
            End If
    
        End With
    
        '<EhFooter>
        Exit Sub

Effect_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mEffect.Effect_Loop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

