Attribute VB_Name = "mSkins"
Option Explicit

Public Type tSkinsData
    ObjIndex As Integer
End Type

Public SkinsLast As Integer
Public ListSkins() As tSkinsData

' #












Public Function Skins_SearchBuyed(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    
    Dim A As Long
    
    With UserList(UserIndex)
        If .Skins.Last > 0 Then
            For A = 1 To .Skins.Last
                If .Skins.ObjIndex(A) = ObjIndex Then
                    Skins_SearchBuyed = True
                    Exit Function
                End If
            Next A
        End If
    End With
End Function

' # Al ingresar al juego recorre la lista de skins existentes y genera la lista de skins permitidas sobre el personaje.
Public Sub Skins_SetChar(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Dim Buyed As Boolean
        Dim A As Long

        For A = 1 To SkinsLast
           ' Buyed = Skins_SearchBuyed(UserIndex)
            
            
            
        Next A
    
    End With
End Sub


Private Function Skins_SearchSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo Skins_SearchSlot_Err
        '</EhHeader>
        Dim A As Long
    
100     With UserList(UserIndex)
    
102         For A = 1 To MAX_INVENTORY_SKINS
        
104             If .Skins.ObjIndex(A) = ObjIndex Then
106                 Skins_SearchSlot = A
                    Exit Function
                End If
108         Next A
    
        End With

        '<EhFooter>
        Exit Function

Skins_SearchSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSkins.Skins_SearchSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function Skins_FreeSlot(ByVal UserIndex As Integer) As Integer

        '<EhHeader>
        On Error GoTo Skins_FreeSlot_Err

        '</EhHeader>
        Dim A As Long
    
100     With UserList(UserIndex)
    
102         For A = 1 To MAX_INVENTORY_SKINS
        
104             If .Skins.ObjIndex(A) = 0 Then
106                 Skins_FreeSlot = A
                    Exit Function

                End If

112         Next A
    
        End With

        '<EhFooter>
        Exit Function

Skins_FreeSlot_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mSkins.Skins_FreeSlot " & "at line " & Erl

        

        '</EhFooter>
End Function

Public Sub Skins_Desequipar(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo Skins_AddNew_Err
        '</EhHeader>
    ' Invalid Obj
    If ObjIndex <= 0 Or ObjIndex > NumObjDatas Then Exit Sub
            
    ' # Invalid Skin
    If ObjData(ObjIndex).Skin <= 0 Then Exit Sub
    
    Dim SlotSkin As Integer
    SlotSkin = Skins_SearchSlot(UserIndex, ObjIndex)
            
    If SlotSkin > 0 Then
            
        With UserList(UserIndex)
            Call Skins_SettingData(UserIndex, .Skins.ObjIndex(SlotSkin), True)
            Call WriteConsoleMsg(UserIndex, "Has desequipado la skin.", FontTypeNames.FONTTYPE_INFOGREEN)
        
        End With
    Else
        ' Error Hack
    End If

    '<EhFooter>
    Exit Sub

Skins_AddNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSkins.Skins_Desequipar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub Skins_AddNew(ByVal UserIndex As Integer, _
                        ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo Skins_AddNew_Err
        '</EhHeader>

        Dim SlotSkin    As Integer
        
        ' Invalid Obj
        If ObjIndex <= 0 Or ObjIndex > NumObjDatas Then Exit Sub
            
        ' # Invalid Skin
110     If ObjData(ObjIndex).Skin <= 0 Then Exit Sub

        ' # No habilitado :c
        If ConfigServer.ModoSkins = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡El sistema de skins está desactivado!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
        
100     With UserList(UserIndex)
        
104         If .Skins.Last = MAX_INVENTORY_SKINS Then
106             Call WriteConsoleMsg(UserIndex, "Tu personaje no permite equipar más skins.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            SlotSkin = Skins_SearchSlot(UserIndex, ObjIndex)
            
            
            If SlotSkin > 0 Then
            
                ' # Chequea si no se actualizó la permanencia de la misma.
                If Not Skins_CheckBuy_Guild(UserIndex, ObjIndex) Then
                    Call Skins_SettingData(UserIndex, .Skins.ObjIndex(SlotSkin), True)
                    Exit Sub
                End If
                
                ' # Setea la skin en las equipadas y la comienza a utilizar
                Call Skins_SettingData(UserIndex, .Skins.ObjIndex(SlotSkin), False)
                
                Call WriteConsoleMsg(UserIndex, "Skin de tu personaje actualizado.", FontTypeNames.FONTTYPE_INFOGREEN)
            Else
                If .Stats.Gld < ObjData(ObjIndex).valor Then Exit Sub
                If .Stats.Eldhir < ObjData(ObjIndex).ValorEldhir Then Exit Sub
                
                
                ' # Chequea si no se actualizó la permanencia de la misma.
                If Not Skins_CheckBuy_Guild(UserIndex, ObjIndex) Then
                    Exit Sub
                End If
                
                ' La intenta comprar
                SlotSkin = Skins_FreeSlot(UserIndex)
                
                If SlotSkin > 0 Then
                    .Stats.Gld = .Stats.Gld - ObjData(ObjIndex).valor
                    .Stats.Eldhir = .Stats.Eldhir - ObjData(ObjIndex).ValorEldhir
                    
                    .Skins.ObjIndex(SlotSkin) = ObjIndex
                    .Skins.Last = .Skins.Last + 1
                    
                    Call WriteUpdateGold(UserIndex)
                    Call WriteUpdateDsp(UserIndex)
                    
                    Call WriteConsoleMsg(UserIndex, "¡Felicitaciones por tu nueva adquisición!", FontTypeNames.FONTTYPE_INFOGREEN)
                    Call Logs_Security(eSecurity, eShop, "Usuario: " & .Account.Email & " Nick: " & .Name & " ORO: " & ObjData(ObjIndex).valor & " DSP: " & ObjData(ObjIndex).ValorEldhir & " ITEM: " & ObjData(ObjIndex).Name & "(" & ObjIndex & ")")
                    
                    ' # la comienza a utilizar
                    Call Skins_SettingData(UserIndex, .Skins.ObjIndex(SlotSkin), False)
                    WriteUpdateDataSkin UserIndex, .Skins.Last
                End If
                
            End If

        End With

        '<EhFooter>
        Exit Sub

Skins_AddNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSkins.Skins_AddNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Chequea si puede comprar la skin
Public Function Skins_CheckBuy_Guild(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    With UserList(UserIndex)
        If ObjData(ObjIndex).GuildLvl > 0 And .GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Esta skin es solo para Clan!", FontTypeNames.FONTTYPE_INFORED)
            Exit Function
        End If
        
        If GuildsInfo(.GuildIndex).Lvl < ObjData(ObjIndex).GuildLvl Then
            Call WriteConsoleMsg(UserIndex, "Esta skin requiere Nivel: " & ObjData(ObjIndex).GuildLvl & ".", FontTypeNames.FONTTYPE_INFORED)
            Exit Function
        End If
    
    End With
    
    Skins_CheckBuy_Guild = True
    
End Function
' # Chequea si el usuario tiene una skin de clan y tiene que sacarsela
Public Function Skins_CheckGuild(ByVal UserIndex As Integer, ByVal Killed As Boolean) As Boolean
    With UserList(UserIndex)
        ' # Esta utilizando una armadura de skin
        If .Skins.ArmourIndex > 0 Then
            ' No tiene clan.
            If .GuildIndex = 0 Then
                If Killed Then Call Skins_SettingData(UserIndex, .Skins.ArmourIndex, True)
                Exit Function
            End If
            
            If ObjData(.Skins.ArmourIndex).GuildLvl > 0 Then
                If GuildsInfo(.GuildIndex).Lvl < ObjData(.Skins.ArmourIndex).GuildLvl Then
                    If Killed Then Call Skins_SettingData(UserIndex, .Skins.ArmourIndex, True)
                    Exit Function
                End If
            End If
        End If
    End With
    
    Skins_CheckGuild = True
    
End Function
' Setea el Skin en el Slot correspondiente
Public Sub Skins_SettingData(ByVal UserIndex As Integer, _
                              ByVal ObjIndexSkin As Integer, _
                              ByVal Killed As Boolean)
        '<EhHeader>
        On Error GoTo Skins_SettingData_Err
        '</EhHeader>
        
        
        Dim Obj As ObjData
        
        
        Obj = ObjData(ObjIndexSkin)
        
100     With UserList(UserIndex)
    
102         Select Case Obj.OBJType
    
                Case eOBJType.otarmadura
                    If Killed Then
                        .Skins.ArmourIndex = 0
                    Else
                        .Skins.ArmourIndex = ObjIndexSkin
                        
                        If Obj.AuraIndex(1) > 0 Then
                            .OrigChar.AuraIndex(1) = Obj.AuraIndex(1)
                            .Char.AuraIndex(1) = .OrigChar.AuraIndex(1)
                        End If
                        
                    End If
                    
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        .Char.Body = GetArmourAnim(UserIndex, IIf(Killed, .Invent.ArmourEqpObjIndex, ObjIndexSkin))
                    End If
                        
                    .OrigChar.Body = .Char.Body

116             Case eOBJType.otWeapon
                    If Obj.Apuñala = 1 Then
                        .Skins.WeaponDagaIndex = IIf(Killed, 0, ObjIndexSkin)
                        
                    ElseIf Obj.proyectil > 0 Then
                        .Skins.WeaponArcoIndex = IIf(Killed, 0, ObjIndexSkin)
                    
                    Else
                        ' Báculos-Espadas
                        .Skins.WeaponIndex = IIf(Killed, 0, ObjIndexSkin)
                        
                    End If
                    
                    If Obj.AuraIndex(2) > 0 Then
                        .OrigChar.AuraIndex(2) = Obj.AuraIndex(2)
                        .Char.AuraIndex(2) = .OrigChar.AuraIndex(2)
                    End If
                        
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, IIf(Killed, .Invent.WeaponEqpObjIndex, ObjIndexSkin))
                    End If
                        
                    .OrigChar.WeaponAnim = .Char.WeaponAnim

130             Case eOBJType.otescudo
                    If Killed Then
                        .Skins.ShieldIndex = 0
                    Else
                        .Skins.ShieldIndex = ObjIndexSkin
                        
                        If Obj.AuraIndex(4) > 0 Then
                            .OrigChar.AuraIndex(4) = Obj.AuraIndex(4)
                            .Char.AuraIndex(4) = .OrigChar.AuraIndex(4)
                        End If
                        
                    End If
                    
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        .Char.ShieldAnim = GetShieldAnim(UserIndex, IIf(Killed, .Invent.EscudoEqpObjIndex, ObjIndexSkin))
                    End If
                        
                    .OrigChar.ShieldAnim = .Char.ShieldAnim

144             Case eOBJType.otcasco
                    If Killed Then
                        .Skins.HelmIndex = 0
                    Else
                        .Skins.HelmIndex = ObjIndexSkin
                        
                        If Obj.AuraIndex(3) > 0 Then
                            .OrigChar.AuraIndex(3) = Obj.AuraIndex(3)
                            .Char.AuraIndex(3) = .OrigChar.AuraIndex(3)
                        End If
                        
                    End If
                    
                    If .Invent.CascoEqpObjIndex > 0 Then
                        .Char.CascoAnim = GetHelmAnim(UserIndex, IIf(Killed, .Invent.CascoEqpObjIndex, ObjIndexSkin))
                    End If
                        
                    .OrigChar.CascoAnim = .Char.CascoAnim

            End Select
              
158         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
            Call WriteUpdateDataSkin(UserIndex, 0)
        End With
  
        '<EhFooter>
        Exit Sub

Skins_SettingData_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSkins.Skins_SettingData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



