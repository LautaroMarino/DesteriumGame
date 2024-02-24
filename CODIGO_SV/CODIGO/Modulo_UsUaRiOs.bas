Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 11/03/2010
        '11/03/2010: ZaMa - Ahora no te vuelve cirminal por matar un atacable
        '***************************************************
        '<EhHeader>
        On Error GoTo ActStats_Err
        '</EhHeader>

        Dim DaExp       As Integer

        Dim EraCriminal As Boolean
    
100     DaExp = CInt(UserList(VictimIndex).Stats.Elv) * 2
    
102     With UserList(AttackerIndex)
104         .Stats.Exp = .Stats.Exp + DaExp

106         If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
108         Call CheckUserLevel(AttackerIndex)
        
110         If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
        
                ' Es legal matarlo si estaba en atacable
112             If UserList(VictimIndex).flags.AtacablePor <> AttackerIndex Then
114                 EraCriminal = Escriminal(AttackerIndex)
                
116                 With .Reputacion

118                     If Not Escriminal(VictimIndex) Then
120                         .AsesinoRep = .AsesinoRep + vlASESINO * 2

122                         If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
124                         .BurguesRep = 0
126                         .NobleRep = 0
128                         .PlebeRep = 0
                        Else
130                         .NobleRep = .NobleRep + vlNoble

132                         If .NobleRep > MAXREP Then .NobleRep = MAXREP
                        End If

                    End With
                

136                 If EraCriminal <> Escriminal(AttackerIndex) Then
138                     Call RefreshCharStatus(AttackerIndex)
                    End If
                
                End If
            End If
        
            'Lo mata
140         Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, VictimIndex, DaExp)
142         Call WriteMultiMessage(VictimIndex, eMessages.UserKill, AttackerIndex)
144         Call FlushBuffer(VictimIndex)
        
            'Log
            'Call Logs_Security(eSecurity, eAntiFrags, .Name & " con IP: " & .Ip & " Y Cuenta: " & .Account.Email & " asesino a " & UserList(VictimIndex).Name & " con IP: " & UserList(VictimIndex).Ip & " Y Cuenta: " & UserList(VictimIndex).Account.Email)
            'Call Logs_User(.Name, eUser, eKill, .Name & " con IP: " & .Ip & " Y Cuenta: " & .Account.Email & " asesino a " & UserList(VictimIndex).Name & " con IP: " & UserList(VictimIndex).Ip & " Y Cuenta: " & UserList(VictimIndex).Account.Email)
        End With

        '<EhFooter>
        Exit Sub

ActStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.ActStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo RevivirUsuario_Err
        '</EhHeader>

100     With UserList(UserIndex)
102         .flags.Muerto = 0
104         .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion) * 5
        
106         If .Stats.MinHp > .Stats.MaxHp Then
108             .Stats.MinHp = .Stats.MaxHp
            End If
        
110         If .flags.Navegando = 1 Then
112             Call ToggleBoatBody(UserIndex)
            Else
114             Call DarCuerpoDesnudo(UserIndex)
            
116             .Char.Head = .OrigChar.Head
            End If
        
118         If .flags.Traveling Then
120             Call EndTravel(UserIndex, True)
            End If
        
122         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
124         Call WriteUpdateUserStats(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

RevivirUsuario_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.RevivirUsuario " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ToggleBoatBody(ByVal UserIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/07/2010
        'Gives boat body depending on user alignment.
        '25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
        '***************************************************
        '<EhHeader>
        On Error GoTo ToggleBoatBody_Err
        '</EhHeader>

        Dim Ropaje        As Integer
        Dim EsFaccionario As Boolean
        Dim NewBody       As Integer
    
100     With UserList(UserIndex)
 
102         .Char.Head = 0
104         If .Invent.BarcoObjIndex = 0 Then Exit Sub
        
106         Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
            
            If Ropaje = 0 Then
            
                ' Criminales y caos
108             If Escriminal(UserIndex) Then
                
110                 EsFaccionario = esCaos(UserIndex)
                
112                 Select Case Ropaje
                        Case iBarca
114                         If EsFaccionario Then
116                             NewBody = iBarcaCaos
                            Else
118                             NewBody = iBarcaPk
                            End If
                    
120                     Case iGalera
122                         If EsFaccionario Then
124                             NewBody = iGaleraCaos
                            Else
126                             NewBody = iGaleraPk
                            End If
                        
128                     Case iGaleon
130                         If EsFaccionario Then
132                             NewBody = iGaleonCaos
                            Else
134                             NewBody = iGaleonPk
                            End If
                    End Select
            
                    ' Ciudas y Armadas
                Else
                
136                 EsFaccionario = esArmada(UserIndex)
                
138                 Select Case Ropaje
                        Case iBarca
140                         If EsFaccionario Then
142                             NewBody = iBarcaReal
                            Else
144                             NewBody = iBarcaCiuda
                            End If
                        
146                     Case iGalera
148                         If EsFaccionario Then
150                             NewBody = iGaleraReal
                            Else
152                             NewBody = iGaleraCiuda
                            End If
                            
154                     Case iGaleon
156                         If EsFaccionario Then
158                             NewBody = iGaleonReal
                            Else
160                             NewBody = iGaleonCiuda
                            End If
                    End Select
                
                End If
                
            Else
                NewBody = Ropaje
            End If
            
162         .Char.Body = NewBody
164         .Char.ShieldAnim = NingunEscudo
166         .Char.WeaponAnim = NingunArma
168         .Char.CascoAnim = NingunCasco

                
              Dim A As Long
              
              For A = 1 To MAX_AURAS
170             .Char.AuraIndex(A) = NingunAura
              Next A
              
        End With

        '<EhFooter>
        Exit Sub

ToggleBoatBody_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.ToggleBoatBody " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Sub ChangeUserChar(ByVal UserIndex As Integer, _
                          ByVal Body As Integer, _
                          ByVal Head As Integer, _
                          ByVal Heading As Byte, _
                          ByVal Arma As Integer, _
                          ByVal Escudo As Integer, _
                          ByVal Casco As Integer, _
                          ByRef AuraIndex() As Byte)
        '<EhHeader>
        On Error GoTo ChangeUserChar_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
100     With UserList(UserIndex).Char
102         .Body = Body
104         .Head = Head
106         .Heading = Heading
108         .WeaponAnim = Arma
110         .ShieldAnim = Escudo
112         .CascoAnim = Casco
              
            
              
              Dim A As Long
              For A = 1 To MAX_AURAS
114             .AuraIndex(A) = AuraIndex(A)
             Next A
             
116        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, 0, Head, Heading, .charindex, Arma, Escudo, .FX, .loops, Casco, AuraIndex, UserList(UserIndex).flags.ModoStream, False, UserList(UserIndex).flags.Navegando))
        End With

        '<EhFooter>
        Exit Sub

ChangeUserChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.ChangeUserChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Function GetArmourAnim_Bot(ByVal Slot As Long, ByVal ObjIndex As Integer) As Integer

        '<EhHeader>
        On Error GoTo GetArmourAnim_Bot_Err

        '</EhHeader>

        '***************************************************
        '
        '
        '
        '***************************************************
        Dim Tmp          As Integer

        Dim SkinSelected As Integer
        
100     With BotIntelligence(Slot)
            
            
          ' If .Skins.Armour > 0 Then
                'ObjIndex = .Skins.Armour
           ' End If

102         Tmp = ObjData(ObjIndex).RopajeEnano
                
104         If Tmp > 0 Then
106             If .Raze = eRaza.Enano Or .Raze = eRaza.Gnomo Then
108                 GetArmourAnim_Bot = Tmp

                    Exit Function

                End If

            End If
        
110         GetArmourAnim_Bot = ObjData(ObjIndex).Ropaje

        End With

        '<EhFooter>
        Exit Function

GetArmourAnim_Bot_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetArmourAnim " & "at line " & Erl

        

        '</EhFooter>
End Function
Public Function GetArmourAnim(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer) As Integer

        '<EhHeader>
        On Error GoTo GetArmourAnim_Err

        '</EhHeader>

        '***************************************************
        '
        '
        '
        '***************************************************
        Dim Tmp          As Integer

        Dim SkinSelected As Integer
        
100     With UserList(UserIndex)

            If .Skins.ArmourIndex > 0 Then
                ObjIndex = .Skins.ArmourIndex
            End If

102         Tmp = ObjData(ObjIndex).RopajeEnano
                
104         If Tmp > 0 Then
106             If .Raza = eRaza.Enano Or .Raza = eRaza.Gnomo Then
108                 GetArmourAnim = Tmp

                    Exit Function

                End If

            End If
        
110         GetArmourAnim = ObjData(ObjIndex).Ropaje

        End With

        '<EhFooter>
        Exit Function

GetArmourAnim_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetArmourAnim " & "at line " & Erl

        

        '</EhFooter>
End Function

Public Function GetShieldAnim(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo GetShieldAnim_Err
        '</EhHeader>
100     With UserList(UserIndex)
102         If .Skins.ShieldIndex > 0 Then
104             ObjIndex = .Skins.ShieldIndex
            End If
        
106          GetShieldAnim = ObjData(ObjIndex).ShieldAnim
        End With
    
        '<EhFooter>
        Exit Function

GetShieldAnim_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.GetShieldAnim " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetHelmAnim(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo GetHelmAnim_Err
        '</EhHeader>
100     With UserList(UserIndex)
102         If .Skins.HelmIndex > 0 Then
104             ObjIndex = .Skins.HelmIndex
            End If
        
106          GetHelmAnim = ObjData(ObjIndex).CascoAnim
        End With
    
        '<EhFooter>
        Exit Function

GetHelmAnim_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.GetHelmAnim " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Public Function GetWeaponAnim(ByVal UserIndex As Integer, _
                              ByVal UserRaza As Byte, _
                              ByVal ObjIndex As Integer) As Integer

        '<EhHeader>
        On Error GoTo GetWeaponAnim_Err

        '</EhHeader>

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 03/29/10
        '
        '***************************************************
        Dim Tmp As Integer

        With UserList(UserIndex)

            If ObjData(ObjIndex).Apuñala = 1 Then
                If .Skins.WeaponDagaIndex > 0 Then
                    ObjIndex = .Skins.WeaponDagaIndex
                End If
            ElseIf ObjData(ObjIndex).proyectil > 0 Then
                If .Skins.WeaponArcoIndex > 0 Then
                    ObjIndex = .Skins.WeaponArcoIndex
                End If
            Else
                If .Skins.WeaponIndex > 0 Then
                    ObjIndex = .Skins.WeaponIndex
                End If
            End If
            
        End With

102     Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
104     If Tmp > 0 Then
106         If UserRaza = eRaza.Enano Or UserRaza = eRaza.Gnomo Then
108             GetWeaponAnim = Tmp

                Exit Function

            End If

        End If
        
110     GetWeaponAnim = ObjData(ObjIndex).WeaponAnim

        '<EhFooter>
        Exit Function

GetWeaponAnim_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetWeaponAnim " & "at line " & Erl
        
        '</EhFooter>
End Function

Public Function GetWeaponAnimBot(ByVal Raza As Byte, ByVal ObjIndex As Integer) As Integer

        '<EhHeader>
        On Error GoTo GetWeaponAnimBot_Err

        '</EhHeader>

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 03/29/10
        '
        '***************************************************
        Dim Tmp As Integer

102     Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
104     If Tmp > 0 Then
106         If Raza = eRaza.Enano Or Raza = eRaza.Gnomo Then
108             GetWeaponAnimBot = Tmp

                Exit Function

            End If

        End If
        
110     GetWeaponAnimBot = ObjData(ObjIndex).WeaponAnim

        '<EhFooter>
        Exit Function

GetWeaponAnimBot_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetWeaponAnim " & "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)
        '*************************************************
        'Author: Unknown
        'Last modified: 08/01/2009
        '08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
        '*************************************************
        '<EhHeader>
        On Error GoTo EraseUserChar_Err
        '</EhHeader>


100     With UserList(UserIndex)


102         CharList(.Char.charindex) = 0
        
104         If .Char.charindex > 0 And .Char.charindex <= LastChar Then
106            CharList(.Char.charindex) = 0
            
108             If .Char.charindex = LastChar Then
110                 Do Until CharList(LastChar) > 0
112                     LastChar = LastChar - 1
114                     If LastChar <= 1 Then Exit Do
                    Loop
                End If
            End If
  
  
                
116         Call ModAreas.DeleteEntity(UserIndex, ENTITY_TYPE_PLAYER)
        
118         If MapaValido(.Pos.Map) Then
120             MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
            End If
                
                
                
122         .Char.charindex = 0
        End With
    
124     NumChars = NumChars - 1



        '<EhFooter>
        Exit Sub

EraseUserChar_Err:
            Dim UserName  As String
    Dim charindex As Integer
    
    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        charindex = UserList(UserIndex).Char.charindex
    End If

    Call LogError("Error en EraseUserchar " & Err.number & ": " & Err.description & ". User: " & UserName & "(UI: " & UserIndex & " - CI: " & charindex & ") en Linea: " & Erl)
        
        '</EhFooter>
End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo RefreshCharStatus_Err
        '</EhHeader>

        '*************************************************
        'Author: Tararira
        'Last modified: 04/07/2009
        'Refreshes the status and tag of UserIndex.
        '04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
        '*************************************************
        Dim ClanTag   As String

        Dim NickColor As Byte
    
100     With UserList(UserIndex)

102         If .GuildIndex > 0 Then
104             ClanTag = GuildsInfo(.GuildIndex).Name
106             ClanTag = " <" & ClanTag & ">"
            End If
        
108         NickColor = GetNickColor(UserIndex)
        
110         If .ShowName Then
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag))
            Else
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, vbNullString))
            End If
        
            'Si esta navengando, se cambia la barca.
116         If .flags.Navegando Then
118             If .flags.Muerto = 1 Then
120                 .Char.Body = iFragataFantasmal
                Else
122                 Call ToggleBoatBody(UserIndex)
                End If
            
124             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

RefreshCharStatus_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.RefreshCharStatus " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function GetNickColor(ByVal UserIndex As Integer) As Byte
        '*************************************************
        'Author: ZaMa
        'Last modified: 15/01/2010
        '
        '*************************************************
        '<EhHeader>
        On Error GoTo GetNickColor_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
102         If Escriminal(UserIndex) Then
104             GetNickColor = eNickColor.ieCriminal
106         ElseIf Not Escriminal(UserIndex) Then
108             GetNickColor = eNickColor.ieCiudadano
            End If
        
110         If .Faction.Status = r_Armada Then
112             GetNickColor = eNickColor.ieArmada
            End If
        
114         If .Faction.Status = r_Caos Then
116             GetNickColor = eNickColor.ieCAOS
            End If
        
118         If .flags.FightTeam = 1 Then
120             GetNickColor = eNickColor.ieCriminal
122         ElseIf .flags.FightTeam = 2 Then
124             GetNickColor = eNickColor.ieCiudadano
            End If
        
            'If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable
        
126         If .Counters.Shield Then
128             GetNickColor = eNickColor.ieShield
            End If
        
130         If Power.UserIndex = UserIndex Then GetNickColor = eNickColor.ieAtacable

        
        End With
    
        '<EhFooter>
        Exit Function

GetNickColor_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.GetNickColor " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function MakeUserChar(ByVal toMap As Boolean, _
                        ByVal sndIndex As Integer, _
                        ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        Optional ByVal ButIndex As Boolean = False, _
                        Optional ByVal IsInvi As Boolean = False) As Boolean
    '*************************************************
    'Author: Unknown
    'Last modified: 15/01/2010
    '23/07/2009: Budi - Ahora se envía el nick
    '15/01/2010: ZaMa - Ahora se envia el color del nick.
    '*************************************************

    On Error GoTo ErrHandler

    Dim charindex  As Integer
    Dim ClanTag    As String
    Dim NickColor  As Byte
    Dim UserName   As String
    Dim Privileges As Byte
    
    With UserList(UserIndex)
    
        If InMapBounds(Map, X, Y) Then

            'If needed make a new character in list
            If .Char.charindex = 0 Then
                charindex = NextOpenCharIndex
                .Char.charindex = charindex
                CharList(charindex) = UserIndex
            End If
            
            'Place character on map if needed
            If toMap Then MapData(Map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = GuildsInfo(.GuildIndex).Name
                End If
                
                NickColor = GetNickColor(UserIndex)
                Privileges = .flags.Privilegios
                
                'Preparo el nick
                If .ShowName Then
                    UserName = .secName
                    
                    If EsSemiDios(UserName) Then
                        UserName = UserName & " " & TAG_GAME_MASTER
                    End If
                    
                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else

                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User) Then
                            If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                        Else

                            If (.flags.Invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else

                                If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                            End If
                        End If
                    End If
                End If
                
                Call WriteCharacterCreate(sndIndex, .Char.Body, 0, .Char.Head, .Char.Heading, .Char.charindex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, UserName, NickColor, Privileges, .Char.AuraIndex, .Char.speeding, False)
                
                If IsInvi Then
                    'Actualizamos las áreas de ser necesario
                    'Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)
                End If
            Else
                ' Me lo mando a mi mismo
                Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y)
                
                ' Se lo mando a los demas
                Call ModAreas.CreateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, ModAreas.DEFAULT_ENTITY_WIDTH, ModAreas.DEFAULT_ENTITY_HEIGHT)
            End If
        End If

    End With

    MakeUserChar = True
    
    Exit Function

ErrHandler:
    Dim UserErrName As String
    Dim UserMap     As Integer
    If UserIndex > 0 Then
        UserErrName = UserList(UserIndex).Name
        UserMap = UserList(UserIndex).Pos.Map
    End If
    
    Dim sError As String
    sError = "MakeUserChar: num: " & Err.number & " desc: " & Err.description & ".User: " & UserErrName & "(" & UserIndex & "). UserMap: " & UserMap & ". Coor: " & Map & "," & X & "," & Y & ". toMap: " & toMap & ". sndIndex: " & sndIndex & ". CharIndex: " & charindex & ". ButIndex: " & ButIndex
    
    '
    Call CloseSocket(UserIndex)
    
    'Para ver si clona..
    sError = sError & ". MapUserIndex: " & MapData(Map, X, Y).UserIndex
    Call LogError(sError)
End Function

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo CheckUserLevel_Err

        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 08/04/2011
        'Chequea que el usuario no halla alcanzado el siguiente nivel,
        'de lo contrario le da la vida, mana, etc, correspodiente.
        '07/08/2006 Integer - Modificacion de los valores
        '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
        '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
        '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
        '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
        '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
        '12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
        '02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
        '11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
        '02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
        '08/04/2011: Amraphen - Arreglada la distribución de probabilidades para la vida en el caso de promedio entero.
        '*************************************************
        Dim Pts              As Integer

        Dim WasNewbie        As Boolean

        Dim promedio         As Double

        Dim aux              As Integer

        Dim DistVida(1 To 5) As Integer

        Dim GI               As Integer 'Guild Index
    
        Dim aumentoHp        As Integer

        Dim AumentoMana      As Integer

        Dim AumentoSta       As Integer

        Dim AumentoHit       As Integer
    
        Dim pasoDeNivel      As Boolean
    
100     WasNewbie = EsNewbie(UserIndex)
    
102     With UserList(UserIndex)

            'Checkea si alcanzó el máximo nivel
108         If .Stats.Elv >= STAT_MAXELV Then
110             .Stats.Exp = 0
112             .Stats.Elu = 0
                
                Exit Sub

            End If

104         Do While .Stats.Exp >= .Stats.Elu And .Stats.Elv < STAT_MAXELV
            
106             pasoDeNivel = True

114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_NIVEL, .Pos.X, .Pos.Y, .Char.charindex))
116             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, FXIDs.FX_LEVEL, 0))
            
118             If .Stats.Elv = 1 Then
120                 Pts = 10
                Else
                    'For multiple levels being rised at once
122                 Pts = Pts + 5

                End If
                
                Dim LastMap As Integer
                
                If MapInfo(.Pos.Map).LvlMax > 0 Then
124                 If .Stats.Elv >= MapInfo(.Pos.Map).LvlMax Then
128                     Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
                        LastMap = .Pos.Map

                    End If

                End If
            
132             .Stats.Elv = .Stats.Elv + 1
                    
                If .Stats.Elv = 35 Then
                    If .Hogar = eCiudad.cEsperanza Then
                        .Hogar = eCiudad.cUllathorpe
                        Call WriteConsoleMsg(UserIndex, "Tu nuevo hogar pasó a ser la Ciudad de Ullathorpe.", FontTypeNames.FONTTYPE_USERGOLD)

                    End If

                End If
                    
                #If Classic = 1 Then
                    
                    'Esta haciendo la mision newbie. Pasamos a la siguiente
                    If .Stats.Elv = LimiteNewbie + 1 Then
                        If .QuestStats(1).QuestIndex > 0 Then
                            .QuestStats(1).QuestIndex = 0
                            Call Quest_SetUser(UserIndex, 2)
                        End If
                        
                    End If
                        
                #End If
                
                If .Stats.Elv >= 35 Then
                    ' # Envia un mensaje a discord
                    Dim TextDiscord As String
                    TextDiscord = "El personaje **'" & .Name & "'** pasó a **nivel " & .Stats.Elv & "**."
                    WriteMessageDiscord CHANNEL_LEVEL, TextDiscord
                End If
                   
134             If .Stats.Elv = STAT_MAXELV Then
136                 .Stats.SkillPts = .Stats.SkillPts + 5
138                 Call WriteLevelUp(UserIndex, .Stats.SkillPts)
140                 Call WriteConsoleMsg(UserIndex, "Has ganado: " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
142                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje " & .Name & " ha alcanzado el nivel máximo.", FontTypeNames.FONTTYPE_INFOGREEN))

                    'Call RankUser_AddPoint(UserIndex, 100)
                End If
            
144             .Stats.Exp = .Stats.Exp - .Stats.Elu
146             .Stats.Elu = EluUser(.Stats.Elv)
148             RecompensaPorNivel UserIndex
            
150             .Stats.MinHp = .Stats.MaxHp
                
                
                
                
            Loop
    
152         If pasoDeNivel Then

                'Send all gained skill points at once (if any)
154             If Pts > 0 Then
156                 Call WriteLevelUp(UserIndex, .Stats.SkillPts + Pts)
                  
158                 .Stats.SkillPts = .Stats.SkillPts + Pts
160                 Call WriteConsoleMsg(UserIndex, "Has ganado: " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                If LastMap > 0 Then
                     Call WriteConsoleMsg(UserIndex, "Vende los objetos que obtuviste en el dungeon y compra algunas pociones más abajo. Busca algun equipamiento básico y recorre el mundo. Ademas puedes verlo desde el botón de arriba.", FontTypeNames.FONTTYPE_USERGOLD)
                End If
                
                ' Comprueba si debe scar objetos
166             Call QuitarLevelObj(UserIndex)
168             Call WriteUpdateUserStats(UserIndex)
                Call SaveUser(UserList(UserIndex), CharPath & UCase$(.Name) & ".chr")
            Else
                
170             Call WriteUpdateExp(UserIndex)

            End If

        End With

        '<EhFooter>
        Exit Sub

CheckUserLevel_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CheckUserLevel " & "at line " & Erl & " Valor Exp: " & UserList(UserIndex).Stats.Exp & " Level User: " & UserList(UserIndex).Stats.Elv

        '</EhFooter>
End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo PuedeAtravesarAgua_Err
        '</EhHeader>

100     PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 Or UserList(UserIndex).flags.Vuela = 1
        '<EhFooter>
        Exit Function

PuedeAtravesarAgua_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.PuedeAtravesarAgua " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading) As Boolean

        '*************************************************
        'Author: Unknown
        'Last modified: 13/07/2009
        'Moves the char, sending the message to everyone in range.
        '30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
        '28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
        '13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
        '13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
        '*************************************************
        '<EhHeader>
        On Error GoTo MoveUserChar_Err

        '</EhHeader>

        Dim nPos               As WorldPos

        Dim sailing            As Boolean

        Dim CasperIndex        As Integer

        Dim CasperHeading      As eHeading

        Dim isAdminInvi        As Boolean

        Dim isZonaOscura       As Boolean

        Dim isZonaOscuraNewPos As Boolean

        Dim UserMoved          As Boolean
 
100     sailing = PuedeAtravesarAgua(UserIndex)
102     nPos = UserList(UserIndex).Pos
104     isZonaOscura = (MapData(nPos.Map, nPos.X, nPos.Y).trigger = eTrigger.zonaOscura)

106     Call HeadtoPos(nHeading, nPos)

108     isZonaOscuraNewPos = (MapData(nPos.Map, nPos.X, nPos.Y).trigger = eTrigger.zonaOscura)
110     isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)

112     If MoveToLegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
114         UserMoved = True

            ' si no estoy solo en el mapa...
116         If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
118             CasperIndex = MapData(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y).UserIndex

                'Si hay un usuario, y paso la validacion, entonces es un casper
120             If CasperIndex > 0 Then

                    ' Los admins invisibles no pueden patear caspers
122                 If Not isAdminInvi Then

124                     With UserList(CasperIndex)
                    
126                         If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
128                             If .flags.SeguroResu = False Then
130                                 .flags.SeguroResu = True
132                                 Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)

                                End If
                             
                            End If
                            
                          '  If .LastHeading > 0 Then
                              '  CasperHeading = .LastHeading
                               ' .LastHeading = 0
                            'Else
                                CasperHeading = InvertHeading(nHeading)

                            'End If

                            '.LastHeading = .Char.Heading
134
136                         Call HeadtoPos(CasperHeading, .Pos)

                            ' Si es un admin invisible, no se avisa a los demas clientes
138                         If Not (.flags.AdminInvisible = 1) Then
140                             ' Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                        
                                'Los valores de visible o invisible están invertidos porque estos flags son del UserIndex, por lo tanto si el UserIndex entra, el casper sale y viceversa :P
142                             If isZonaOscura Then
144                                 If Not isZonaOscuraNewPos Then
146                                     Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                                    End If

                                Else

148                                 If isZonaOscuraNewPos Then
150                                     Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.charindex, False))

                                    End If

                                End If

                            End If

152                         Call WriteForceCharMove(CasperIndex, CasperHeading)

                            'Update map and char
154                         .Char.Heading = CasperHeading
156                         MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = CasperIndex
                        
                            'Actualizamos las áreas de ser necesario
158                         Call ModAreas.UpdateEntity(CasperIndex, ENTITY_TYPE_PLAYER, .Pos)

                        End With

                    End If

                End If

                ' Si es un admin invisible, no se avisa a los demas clientes
                'If Not isAdminInvi Then Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))
            
            End If

            ' Los admins invisibles no pueden patear caspers
160         If (Not isAdminInvi) Or (CasperIndex = 0) Then

162             With UserList(UserIndex)

                    ' Si no hay intercambio de pos con nadie
164                 If CasperIndex = 0 Then
166                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0

                    End If

168                 .Pos = nPos
170                 .Char.Heading = nHeading
172                 MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                
174                 If Extra.IsAreaResu(UserIndex) Then
176                     Call Extra.AutoCurar(UserIndex)

                    End If

178                 If isZonaOscura Then
180                     If Not isZonaOscuraNewPos Then
182                         If (.flags.Invisible Or .flags.Oculto) = 0 Then
184                             Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))

                            End If

                        End If

                    Else

186                     If isZonaOscuraNewPos Then
188                         If (.flags.Invisible Or .flags.Oculto) = 0 Then
190                             Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                            End If

                        End If

                    End If

202                 If .flags.SlotEvent > 0 Then
204                     If Events(.flags.SlotEvent).Modality = Busqueda Then
206                         If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjEvent = 1 Then
208                             Call EventosDS.Busqueda_GetObj(.flags.SlotEvent, .flags.SlotUserEvent)
210                             MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjEvent = 0
212                             EraseObj 10000, .Pos.Map, .Pos.X, .Pos.Y

                            End If

                        End If

                    End If

                    ' // NUEVO
214                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then
216                     Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)

                    End If
                
                    'Actualizamos las áreas de ser necesario
218                 Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)

                End With

            Else
220             Call WritePosUpdate(UserIndex)

            End If

        Else
222         Call WritePosUpdate(UserIndex)

        End If

224     If UserList(UserIndex).Counters.Trabajando Then UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
226     If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
        
228     MoveUserChar = UserMoved
    
        '<EhFooter>
        Exit Function

MoveUserChar_Err:
        Call LogError("Error " & Err.number & " (Linea: " & Erl & ") " & Err.description & " en User: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).IpAddress & " con pos " & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
        Call LogError("Error " & Err.number & " (Linea: " & Erl & ") " & Err.description & " en User: " & UserList(CasperIndex).Name & " con IP: " & UserList(CasperIndex).IpAddress & " con pos " & UserList(CasperIndex).Pos.Map & " X:" & UserList(CasperIndex).Pos.X & " Y:" & UserList(CasperIndex).Pos.Y)

        '</EhFooter>
End Function
Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
        '<EhHeader>
        On Error GoTo InvertHeading_Err
        '</EhHeader>

        '*************************************************
        'Author: ZaMa
        'Last modified: 30/03/2009
        'Returns the heading opposite to the one passed by val.
        '*************************************************
100     Select Case nHeading

            Case eHeading.EAST
102             InvertHeading = WEST

104         Case eHeading.WEST
106             InvertHeading = EAST

108         Case eHeading.SOUTH
110             InvertHeading = NORTH

112         Case eHeading.NORTH
114             InvertHeading = SOUTH
        End Select

        '<EhFooter>
        Exit Function

InvertHeading_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.InvertHeading " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ChangeUserInv_Err
        '</EhHeader>

100     UserList(UserIndex).Invent.Object(Slot) = Object
102     Call WriteChangeInventorySlot(UserIndex, Slot)
        '<EhFooter>
        Exit Sub

ChangeUserInv_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.ChangeUserInv " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function NextOpenCharIndex() As Integer
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo NextOpenCharIndex_Err
        '</EhHeader>

        Dim LoopC As Long
    
100     For LoopC = 1 To MAXCHARS

102         If CharList(LoopC) = 0 Then
104             NextOpenCharIndex = LoopC
106             NumChars = NumChars + 1
            
108             If LoopC > LastChar Then LastChar = LoopC
            
                Exit Function

            End If

110     Next LoopC

        '<EhFooter>
        Exit Function

NextOpenCharIndex_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.NextOpenCharIndex " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub FreeSlot(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 01/10/2012
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo FreeSlot_Err
        '</EhHeader>

100     If UserIndex = LastUser Then

102         Do While (LastUser > 0)

104             If UserList(LastUser).ConnIDValida Then Exit Do
106             LastUser = LastUser - 1
            Loop

        End If

        '<EhFooter>
        Exit Sub

FreeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.FreeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Function PonerPuntos(numero As Long) As String
    
    On Error GoTo PonerPuntos_Err
    

    Dim i     As Integer

    Dim Cifra As String
 
    Cifra = Str(numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)

    For i = 0 To 4

        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos

            End If

        Else

            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos

            End If

            Exit For

        End If

    Next
 
    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
    
    Exit Function

PonerPuntos_Err:
    'Call RegistrarError(err.Number, err.Description, "ModLadder.PonerPuntos", Erl)
    
    
End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 26/05/2011 (Amraphen)
        '26/05/2011: Amraphen - Ahora envía la defensa adicional de la armadura de segunda jerarquía
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserStatsTxt_Err
        '</EhHeader>

        Dim GuildI             As Integer

        Dim ModificadorDefensa As Single 'Por las armaduras de segunda jerarquía.

        Dim Ups                As Single

        Dim UpsSTR             As String
    
100     With UserList(UserIndex)
102         Ups = .Stats.MaxHp - Mod_Balance.getVidaIdeal(.Stats.Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion))
        
104         If Ups > 0 Then
106             UpsSTR = "+" & Ups
108         ElseIf Ups < 0 Then
110             UpsSTR = Ups
            Else
112             UpsSTR = "promedio"

            End If
        
114         Call WriteConsoleMsg(sendIndex, "Personaje: " & .Name & ". " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & " " & UpsSTR, FontTypeNames.FONTTYPE_INFO)
116         Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.Elv & "  EXP: " & PonerPuntos(CLng(.Stats.Exp)) & "/" & PonerPuntos(.Stats.Elu), FontTypeNames.FONTTYPE_INFO)
118         'Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.Clase) & " " & ListaRazas(.Raza), FontTypeNames.FONTTYPE_INFO)
                
              If EsGmPriv(sendIndex) Then
                'Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.Gld & " Eldhires: " & .Stats.Eldhir & " Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.Map, FontTypeNames.FONTTYPE_INFO)
120             'Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & " " & UpsSTR & " Maná: " & .Stats.MinMan & "/" & .Stats.MaxMan & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
              End If
              
            Dim Faction As String
        
122         If .Faction.Status <> r_None Then
124             Faction = InfoFaction(.Faction.Status).Name & " <" & InfoFaction(.Faction.Status).Range(.Faction.Range).Text & ">"
            
126             Call WriteConsoleMsg(sendIndex, "Facción: " & Faction, FontTypeNames.FONTTYPE_INFO)

            End If
        
128         If .Faction.FragsCri > 0 Then Call WriteConsoleMsg(sendIndex, "Criminales Asesinados: " & .Faction.FragsCri, FontTypeNames.FONTTYPE_INFO)
130         If .Faction.FragsCiu > 0 Then Call WriteConsoleMsg(sendIndex, "Ciudadanos Asesinados: " & .Faction.FragsCiu, FontTypeNames.FONTTYPE_INFO)
132         If .flags.Traveling = 1 Then Call WriteConsoleMsg(sendIndex, "Tiempo restante para llegar a tu hogar: " & GetHomeArrivalTime(UserIndex) & " segundos.", FontTypeNames.FONTTYPE_INFO)

136         If .Counters.TimeTelep > 0 Then Call WriteConsoleMsg(sendIndex, "Tiempo para irte del mapa: " & Int(.Counters.TimeTelep / 60) & " minuto/s", FontTypeNames.FONTTYPE_INFO)
        
138         If .Counters.TimeBono > 0 Then Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto gema: " & Int(.Counters.TimeBono / 60) & " minuto/s", FontTypeNames.FONTTYPE_INFO)

140         If .Counters.Pena > 0 Then
142             Call WriteConsoleMsg(sendIndex, "Tiempo restante para salir en libertad: " & .Counters.Pena & " minuto" & IIf(.Counters.Pena = 1, vbNullString, "s"), FontTypeNames.FONTTYPE_INFOGREEN)

            End If
        
144             If .Counters.TimeBonus > 0 Then
146                 If .Counters.TimeBonus < 60 Then
148                     Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto: " & .Counters.TimeBonus & " segundos.", FontTypeNames.FONTTYPE_INFO)
150                 ElseIf .Counters.TimeBonus = 60 Then
152                     Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto: 1 minuto.", FontTypeNames.FONTTYPE_INFO)
                    Else
154                     Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto: " & Int(.Counters.TimeBonus / 60) & " minutos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
156         Call WriteConsoleMsg(sendIndex, "Puntos de Torneo: " & .Stats.Points, FontTypeNames.FONTTYPE_USERGOLD)
158         Call WriteConsoleMsg(sendIndex, "Reputación: " & .Reputacion.promedio & "." & IIf(.Reputacion.promedio < 0, " Paga " & PonerPuntos(5 * Abs(.Reputacion.promedio) * 6) & " Monedas de Oro para ser Ciudadano.", vbNullString), FontTypeNames.FONTTYPE_ANGEL)
         
160         If UserIndex = sendIndex Then
            
            
162             If .Account.Premium > 0 Then
164                 Call WriteConsoleMsg(sendIndex, "Tiempo de Tier: " & .Account.Premium & " restante " & .Account.DatePremium & ".", FontTypeNames.FONTTYPE_USERGOLD)

                End If
            
           

            End If
        
        End With

        '<EhFooter>
        Exit Sub

SendUserStatsTxt_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SendUserStatsTxt " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserInvTxt_Err
        '</EhHeader>


        Dim j As Long
    
100     With UserList(UserIndex)
102         Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
106         For j = 1 To .CurrentInventorySlots

108             If .Invent.Object(j).ObjIndex > 0 Then
110                 Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
                End If

112         Next j

        End With

        '<EhFooter>
        Exit Sub

SendUserInvTxt_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SendUserInvTxt " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserInvTxtFromChar_Err
        '</EhHeader>


        Dim j        As Long

        Dim Charfile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long
    
100     Charfile = CharPath & charName & ".chr"
    
102     If FileExist(Charfile, vbNormal) Then
104         Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(Charfile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
108         For j = 1 To MAX_INVENTORY_SLOTS
110             Tmp = GetVar(Charfile, "Inventory", "Obj" & j)
112             ObjInd = ReadField(1, Tmp, Asc("-"))
114             ObjCant = ReadField(2, Tmp, Asc("-"))

116             If ObjInd > 0 Then
118                 Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
                End If

120         Next j

        Else
122         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

SendUserInvTxtFromChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SendUserInvTxtFromChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserSkillsTxt_Err
        '</EhHeader>


        Dim j As Integer
    
100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
102     For j = 1 To NUMSKILLS
104         Call WriteConsoleMsg(sendIndex, InfoSkill(j).Name & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
106     Next j
    
108     Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
        '<EhFooter>
        Exit Sub

SendUserSkillsTxt_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SendUserSkillsTxt " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, _
                                    ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EsMascotaCiudadano_Err
        '</EhHeader>

100     If Npclist(NpcIndex).MaestroUser > 0 Then
102         EsMascotaCiudadano = Not Escriminal(Npclist(NpcIndex).MaestroUser)

104         If EsMascotaCiudadano Then
106             Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        '<EhFooter>
        Exit Function

EsMascotaCiudadano_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.EsMascotaCiudadano " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo NPCAtacado_Err
        '</EhHeader>

        '**********************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
        '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
        '06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
        '02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
        '**********************************************
        Dim EraCriminal As Boolean
    
        'Guardamos el usuario que ataco el npc.
100     If Npclist(NpcIndex).Movement <> Estatico And Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
102         Npclist(NpcIndex).Target = UserIndex
104         Npclist(NpcIndex).Hostile = 1
106         Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
        End If

        'Npc que estabas atacando.
        Dim LastNpcHit As Integer

108     LastNpcHit = UserList(UserIndex).flags.NPCAtacado
        'Guarda el NPC que estas atacando ahora.
110     UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
        'Revisamos robo de npc.
        'Guarda el primer nick que lo ataca.
112     If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then

            'El que le pegabas antes ya no es tuyo
114         If LastNpcHit <> 0 Then
116             If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
118                 Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
                End If
            End If

120         Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
122     ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then

            'Estas robando NPC
            'El que le pegabas antes ya no es tuyo
124         If LastNpcHit <> 0 Then
126             If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
128                 Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
                End If
            End If
        End If
    
130     If Npclist(NpcIndex).MaestroUser > 0 Then
132         If Npclist(NpcIndex).MaestroUser <> UserIndex Then
134             Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
        End If
    
136     If EsMascotaCiudadano(NpcIndex, UserIndex) Then
138         Call VolverCriminal(UserIndex)
140         Npclist(NpcIndex).Movement = TipoAI.NpcDefensa
142         Npclist(NpcIndex).Hostile = 1
        Else
144         EraCriminal = Escriminal(UserIndex)
        
            'Reputacion
146         If Npclist(NpcIndex).flags.AIAlineacion = 0 Then
148             If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
150                 Call VolverCriminal(UserIndex)
                End If
        
152         ElseIf Npclist(NpcIndex).flags.AIAlineacion = 1 Then
154             UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2

156             If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then UserList(UserIndex).Reputacion.PlebeRep = MAXREP
            End If
        
158         If Npclist(NpcIndex).MaestroUser <> UserIndex Then
                'hacemos que el npc se defienda
160             Npclist(NpcIndex).Movement = TipoAI.NpcDefensa
162             Npclist(NpcIndex).Hostile = 1
            End If
        
164         If EraCriminal And Not Escriminal(UserIndex) Then
166             Call VolverCiudadano(UserIndex)
            End If
        
168         Call AllMascotasAtacanNPC(NpcIndex, UserIndex)
        End If
        
        
        If UserList(UserIndex).GuildIndex > 0 Then
            If Npclist(NpcIndex).CastleIndex > 0 Then
                Call Castle_Attack(Npclist(NpcIndex).CastleIndex, UserList(UserIndex).GuildIndex)
            End If
        End If
        
        '<EhFooter>
        Exit Sub

NPCAtacado_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.NPCAtacado " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo PuedeApuñalar_Err
        '</EhHeader>
    
        Dim WeaponIndex As Integer
     
100     With UserList(UserIndex)
        
102         WeaponIndex = .Invent.WeaponEqpObjIndex
        
104         If WeaponIndex > 0 Then
106             If ObjData(WeaponIndex).Apuñala = 1 Then
108                 PuedeApuñalar = .Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR Or .Clase = eClass.Assasin
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Function

PuedeApuñalar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.PuedeApuñalar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PuedeAcuchillar(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/01/2010 (ZaMa)
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo PuedeAcuchillar_Err
        '</EhHeader>
    
        Dim WeaponIndex As Integer
    
100     With UserList(UserIndex)

102         If .Clase = eClass.Thief Then
        
104             WeaponIndex = .Invent.WeaponEqpObjIndex

106             If WeaponIndex > 0 Then
108                 PuedeAcuchillar = (ObjData(WeaponIndex).Acuchilla = 1)
                End If
            End If

        End With
    
        '<EhFooter>
        Exit Function

PuedeAcuchillar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.PuedeAcuchillar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub SubirSkill(ByVal UserIndex As Integer, _
               ByVal Skill As Integer, _
               ByVal Acerto As Boolean)
        '*************************************************
        'Author: Unknown
        'Last modified: 11/19/2009
        '11/19/2009 Pato - Implement the new system to train the skills.
        '*************************************************
        '<EhHeader>
        On Error GoTo SubirSkill_Err
        '</EhHeader>

        Dim SubeSkill As Boolean
    
100     With UserList(UserIndex)

102         If .flags.Hambre = 0 And .flags.Sed = 0 Then

104             With .Stats

106                 If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
108                 If .UserSkills(Skill) >= LevelSkill(.Elv).LevelValue Then Exit Sub
                      
110                 If Acerto Then
112                     If RandomNumber(1, 100) <= 50 Then SubeSkill = True
                    Else

114                     If RandomNumber(1, 100) <= 20 Then SubeSkill = True
                    End If
                
116                 If SubeSkill Then
118                     .UserSkills(Skill) = .UserSkills(Skill) + 1
120                     Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & InfoSkill(Skill).Name & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                    
122                     .Exp = .Exp + 50

124                     If .Exp > MAXEXP Then .Exp = MAXEXP
                    
126                     Call WriteConsoleMsg(UserIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    
128                     Call WriteUpdateExp(UserIndex)
130                     Call CheckUserLevel(UserIndex)
                        'Call CheckEluSkill(UserIndex, Skill, False)
                    End If

                End With

            End If

        End With

        '<EhFooter>
        Exit Sub

SubirSkill_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SubirSkill " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Public Sub UserDie(ByVal UserIndex As Integer, _
                   Optional ByVal AttackerIndex As Integer = 0)

    '************************************************
    'Author: Uknown
    'Last Modified: 12/01/2010 (ZaMa)
    '04/15/2008: NicoNZ - Ahora se resetea el counter del invi
    '13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
    '27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
    '21/07/2009: Marco - Al morir se desactiva el comercio seguro.
    '16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
    '27/11/2009: Budi - Al morir envia los atributos originales.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
    '************************************************
    On Error GoTo ErrorHandler

    Dim i           As Long

    Dim aN          As Integer
    
    Dim iSoundDeath As Integer
    
    Dim A As Long
    
    
    Dim Time As Long
    
    With UserList(UserIndex)
        
        ' # Masacre en mapas inseguros.
        If MapInfo(.Pos.Map).Pk And .flags.SlotEvent = 0 And .flags.SlotReto = 0 And .flags.SlotFast = 0 Then
            Time = GetTime
            
            If Time - MapInfo(.Pos.Map).DeadTime <= 30000 Then
                MapInfo(.Pos.Map).UsersDead = MapInfo(.Pos.Map).UsersDead + 1
                MapInfo(.Pos.Map).DeadTime = Time
                
                If MapInfo(.Pos.Map).UsersDead > 3 Then
                     WriteMessageDiscord CHANNEL_ONFIRE, "Masacre en **" & MapInfo(.Pos.Map).Name & "**. " & MapInfo(.Pos.Map).UsersDead & _
                     " víctimas caídas en menos de 30 segundos. Players: **" & MapInfo(.Pos.Map).NumUsers & "**"
                End If
           
            Else
                MapInfo(.Pos.Map).UsersDead = 0
                MapInfo(.Pos.Map).DeadTime = 0
            End If
        End If
        
        
        'Sonido
        If .Genero = eGenero.Mujer Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_MUJER_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_MUJER
            End If

        Else

            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE
            End If
        End If
        
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, iSoundDeath)
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
        
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1
        
        .Counters.Trabajando = 0
        
        ' No se activa en arenas
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
        End If
        
        aN = .flags.AtacadoPorNpc

        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
            Npclist(aN).Target = 0
        End If
        
        aN = .flags.NPCAtacado

        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If

        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        Call PerdioNpc(UserIndex, False)
        
        '<<<< Atacable >>>>
        If .flags.AtacablePor > 0 Then
            .flags.AtacablePor = 0
            Call RefreshCharStatus(UserIndex)
        End If
        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
            
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0, .Pos.X, .Pos.Y))
        End If
        
        '<<<< Invisible >>>>
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call SetInvisible(UserIndex, .Char.charindex, False)
        End If
        
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 1
        End If
        
        If MapInfo(.Pos.Map).CaenItems > 0 Then
            If Not EsGm(UserIndex) Then
                If TieneObjetos(PENDIENTE_SACRIFICIO, 1, UserIndex) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_WARP, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 54, 2))
                    Call QuitarObjetos(PENDIENTE_SACRIFICIO, 1, UserIndex)
                
                Else
                        If MapInfo(.Pos.Map).CaenItems = 1 Then
                            Call TirarTodo(UserIndex)
                        End If
                End If
            End If
        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If
        
        ' Desequipamos la montura
        If .Invent.MonturaObjIndex > 0 Then
            
            Call Desequipar(UserIndex, .Invent.MonturaSlot)
            
            If .flags.Montando Then

                .flags.Montando = False
                Call WriteMontateToggle(UserIndex)
            End If
        End If
        
        ' Desequipamos el pendiente de experiencia
        If .Invent.PendientePartyObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.PendientePartySlot)
        End If
        
        ' Desequipamos la reliquia
        If .Invent.ReliquiaObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ReliquiaSlot)
        End If
        
        ' Desequipamos el Objeto mágico (Laudes y Anillos mágicos)
        If .Invent.MagicObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
        End If
        
        'desequipar aura
        If .Invent.AuraEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.AuraEqpSlot)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar anillo magico/laud
        If .Invent.MagicSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)
        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        End If
        

        ' << Restauramos el mimetismo
        If .flags.Mimetizado = 1 Or .flags.Transform = 1 Or .flags.TransformVIP = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
            Next A
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Puede ser atacado por npcs (cuando resucite)
            .flags.Ignorado = False
            .ShowName = True
            
        End If
        
        ' << Restauramos la transformación
        If .flags.Transform = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
            Next A
            
            .Counters.TimeTransform = 0
            .flags.Transform = 0
            .flags.Mimetizado = 0
            ' Puede ser atacado por npcs (cuando resucite)
        End If
        
        ' << Restauramos la transformación VIP
        If .flags.TransformVIP = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex(A)
            Next A
            .flags.TransformVIP = 0
            .flags.Mimetizado = 0
            
            ' Puede ser atacado por npcs (cuando resucite)
            .flags.Ignorado = False
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then

            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i

        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
            .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = NingunArma
            Next A
            Debug.Print .Char.Head
        Else
            .Char.Body = iFragataFantasmal
        End If

        If .MascotaIndex > 0 Then
            Call MuereNpc(.MascotaIndex, 0)
        End If
        
        ' Chequeos del Poder de las Medusas y lo saco al morir
        If Power.UserIndex = UserIndex Then
            If AttackerIndex > 0 Then
                'If Power.Active Then
                    Call Power_Set(AttackerIndex, UserIndex)
                    Call Power_Message
                'Else
                    'Call Power_Set(0, 0)
                'End If
            Else
                Call Power_Set(0, UserIndex)
            End If
        End If
        
        '<< Actualizamos clientes >>
        Call RefreshCharStatus(UserIndex)
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, .Char.AuraIndex)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(UserIndex)
        
        ' Hay que teletransportar?
        Dim mapa As Integer

        mapa = .Pos.Map

        Dim MapaTelep As Integer

        MapaTelep = MapInfo(mapa).OnDeathGoTo.Map
        
        If MapaTelep <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(UserIndex, MapaTelep, MapInfo(mapa).OnDeathGoTo.X, MapInfo(mapa).OnDeathGoTo.Y, True, True)
        End If
        
        ' Retos
        If .flags.SlotReto Then
            Call mRetos.UserdieFight(UserIndex, AttackerIndex, False)
        End If
        
        ' Desafios
        If .flags.Desafiando > 0 Then
            Desafio_UserKill UserIndex
        End If
        
        ' Retos Rapidos
        If .flags.SlotFast > 0 Then
            RetoFast_UserDie UserIndex
        End If
        
        ' Eventos automáticos
        If .flags.SlotEvent > 0 Then
            Call Events_UserDie(UserIndex, AttackerIndex)
        End If
        

    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.number & " Descripción: " & Err.description)
End Sub

Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 13/07/2010
        '13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
        '***************************************************
        '<EhHeader>
        On Error GoTo ContarMuerte_Err
        '</EhHeader>

100     If EsNewbie(Muerto) Then Exit Sub
        
102     With UserList(Atacante)
              'Dim Value As Long
104         'Value = CLng(.Stats.Elv - UserList(Muerto).Stats.Elv)
106        ' If Abs(Value) > 12 Then Exit Sub
             
108         If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
110         If AntiFrags_CheckUser(Atacante, Muerto, 1800) = False Then Exit Sub
        
112         If Not MapInfo(.Pos.Map).FreeAttack Then
114             If Escriminal(Muerto) Then
116                 If .flags.LastCrimMatado <> UserList(Muerto).Name Then
118                     .flags.LastCrimMatado = UserList(Muerto).Name
    
120                     If .Faction.FragsCri < MAXUSERMATADOS Then .Faction.FragsCri = .Faction.FragsCri + 1
                    End If
    
                Else
    
122                 If .flags.LastCiudMatado <> UserList(Muerto).Name Then
124                     .flags.LastCiudMatado = UserList(Muerto).Name
    
126                     If .Faction.FragsCiu < MAXUSERMATADOS Then .Faction.FragsCiu = .Faction.FragsCiu + 1
                    End If
                End If
            End If
        
128         If .Faction.FragsOther < MAXUSERMATADOS Then .Faction.FragsOther = .Faction.FragsOther + 1
        
            'Call RankUser_AddPoint(Atacante, 1)
        End With

        '<EhFooter>
        Exit Sub

ContarMuerte_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.ContarMuerte " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, _
              ByRef nPos As WorldPos, _
              ByRef Obj As Obj, _
              ByRef PuedeAgua As Boolean, _
              ByRef PuedeTierra As Boolean)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/09/2010
    '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
    '18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
    '**************************************************************
    On Error GoTo ErrHandler

    Dim Found As Boolean

    Dim LoopC As Integer

    Dim tX    As Long

    Dim tY    As Long
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, True) Then
        
        If Not HayObjeto(Pos.Map, nPos.X, nPos.Y, Obj.ObjIndex, Obj.Amount) Then
            Found = True
        End If
        
    End If
    
    ' Busca en las demas posiciones, en forma de "rombo"
    If Not Found Then

        While (Not Found) And LoopC <= 16

            If RhombLegalTilePos(Pos, tX, tY, LoopC, Obj.ObjIndex, Obj.Amount, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True
            End If
        
            LoopC = LoopC + 1

        Wend
        
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en Tilelibre. Error: " & Err.number & " - " & Err.description)
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 ByVal FX As Boolean, _
                 Optional ByVal Teletransported As Boolean)
        '<EhHeader>
        On Error GoTo WarpUserChar_Err
        '</EhHeader>

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 11/23/2010
        '15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
        '13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
        '16/09/2010 - ZaMa: No se pierde la visibilidad al cambiar de mapa al estar navegando invisible.
        '11/23/2010 - C4b3z0n: Ahora si no se permite Invi o Ocultar en el mapa al que cambias, te lo saca
        '**************************************************************
        Dim OldMap As Integer

        Dim OldX   As Integer

        Dim OldY   As Integer

        Dim nPos   As WorldPos
    
          If Map = 0 Or X = 0 Or Y = 0 Then
            Call LogError("Cuenta " & UserList(UserIndex).Account.Email & " NICK: " & UserList(UserIndex).Name & "  Map " & Map & " X: " & X & " Y: " & Y)
            Exit Sub
          End If
          
100     With UserList(UserIndex)
            'Quitar el dialogo solo si no es GM.
102         If .flags.AdminInvisible = 0 Then
104             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
            End If
      
106         OldMap = .Pos.Map
108         OldX = .Pos.X
110         OldY = .Pos.Y
              
              
112        ' If OldMap <> Map Then
            

114             If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                
                    Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)
                    Dim WasInvi      As Boolean

                    'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
116                 If MapInfo(Map).InviSinEfecto > 0 And .flags.Invisible = 1 Then
118                     .flags.Invisible = 0
120                     .Counters.Invisibilidad = 0
122                     AhoraVisible = True
124                     WasInvi = True 'si era invi, para el string
                    End If

                    'Chequeo de flags de mapa por ocultar (C4b3z0n)
126                 If MapInfo(Map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
128                     AhoraVisible = True
130                     .flags.Oculto = 0
132                     .Counters.TiempoOculto = 0
                    End If
                
                    ' Chequeo de flags de gran poder
134                 If Power.UserIndex = UserIndex Then
136                     If Not MapInfo(Map).Poder = 1 Then
138                         Call Power_Set(0, UserIndex)
                        End If
                    End If
                
                    'Chequeo de Mimetismo de mapa
140                 If MapInfo(Map).MimetismoSinEfecto > 0 And .flags.Mimetizado = 1 Then
142                     Call Mimetismo_Reset(UserIndex)
                    End If

144                 If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
146                     Call SetInvisible(UserIndex, .Char.charindex, False)

148                     If WasInvi Then 'era invi
150                         Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                        Else 'estaba oculto
152                         Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
            
                Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
154
156             Call WritePlayMusic(UserIndex, val(ReadField(1, MapInfo(Map).Music, 45)))

160             Call WriteChangeMap(UserIndex, Map)
                  'Call WritePosUpdate(UserIndex)
                
                'Update old Map Users
166             MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
168             MapInfo(OldMap).Players.Remove UserIndex
            
                'Update new Map Users
162             MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
164             MapInfo(Map).Players.Add UserIndex

170             If MapInfo(OldMap).NumUsers < 0 Then
172                 MapInfo(OldMap).NumUsers = 0
                End If
        
                'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
                Dim nextMap, previousMap As Boolean

174             nextMap = IIf(distanceToCities(Map).distanceToCity(1) >= 0, True, False)
176             previousMap = IIf(distanceToCities(.Pos.Map).distanceToCity(1) >= 0, True, False)

178             If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                    'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
180             ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
182                 .flags.LastMap = .Pos.Map
184             ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
186                 .flags.LastMap = 0
188             ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
190                 .flags.LastMap = .flags.LastMap
                End If

192             If .flags.Privilegios = PlayerType.User Or .flags.Privilegios = PlayerType.RoyalCouncil Or .flags.Privilegios = PlayerType.ChaosCouncil Then
194                 Call WriteRemoveAllDialogs(UserIndex)
                End If
            
            
          '  Else
196        '     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
198          '   MapData(.Pos.Map, X, Y).UserIndex = UserIndex
            'End If

200         .Pos.X = X
202         .Pos.Y = Y
204         .Pos.Map = Map
                
206         'If OldMap <> Map Then

208             Call MakeUserChar(True, Map, UserIndex, Map, X, Y)
210             Call WriteUserCharIndexInServer(UserIndex)
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterAttackMovement(UserList(UserIndex).Char.charindex), , True)
                      'Actualizamos las áreas de ser necesario
                '   Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)
          '  End If
          
          
            ' // NUEVO
216         If MapData(Map, X, Y).TileExit.Map > 0 Then
218             Call DoTileEvents(UserIndex, Map, X, Y)
            End If

            'Seguis invisible al pasar de mapa
220         If (.flags.Invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            
                ' No si estas navegando
222             If .flags.Navegando = 0 Then
224                 Call SetInvisible(UserIndex, .Char.charindex, True)
                End If
            End If

226         If Teletransported Then
228             If .flags.Traveling = 1 Then
230                 Call EndTravel(UserIndex, True)
                End If
            End If
        
232         If FX And .flags.AdminInvisible = 0 And Not EsAdmin(.Name) Then  'FX
234             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_WARP, X, Y))
236             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FXWARP, 0))
            End If
        
238         If .MascotaIndex Then
240             Call QuitarPet(UserIndex, .MascotaIndex)
                'Call WarpMascota_Map(UserIndex)
            End If

            ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
242         Call IntervaloPermiteSerAtacado(UserIndex, True)
        
            ' Perdes el npc al cambiar de mapa
244         Call PerdioNpc(UserIndex, False)

            ' Automatic toogle navigate
246         If (.flags.Privilegios And (PlayerType.User)) = 0 Then
248             If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
250                 If .flags.Navegando = 0 Then
252                     .flags.Navegando = 1
                        
                        'Tell the client that we are navigating.
254                     Call WriteNavigateToggle(UserIndex)
                    End If

                Else

256                 If .flags.Navegando = 1 Then
258                     .flags.Navegando = 0
                            
                        'Tell the client that we are navigating.
260                     Call WriteNavigateToggle(UserIndex)
                    End If
                End If
            End If

            ' Checking Event Teleports
262         If .flags.SlotEvent > 0 Then
264             If Events(.flags.SlotEvent).Modality = eModalityEvent.Teleports Then
266                 If .Pos.Map = MapEvent.TeleportWin.Map And .Pos.X = MapEvent.TeleportWin.X And .Pos.Y = MapEvent.TeleportWin.Y Then
268                     Events_Teleports_Finish UserIndex
                    End If
                End If
            End If

        
        End With
    
        Exit Sub

WarpUserChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.WarpUserChar (Map: " & Map & " X: " & X & " Y: " & Y & ")" & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WarpMascota_Map(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WarpMascota_Map_Err
        '</EhHeader>

        '************************************************
        'Author: Uknown
        'Last Modified: 26/10/2010
        '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
        '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
        '11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
        '26/10/2010: ZaMa - Ahora las mascotas rapswnean de forma aleatoria.
        '************************************************

        Dim PetTiempoDeVida As Integer

        Dim canWarp         As Boolean

        Dim Index           As Integer

        Dim iMinHP          As Integer
    
        Dim NpcNumber       As Integer

100     With UserList(UserIndex)
102         canWarp = (MapInfo(.Pos.Map).Pk = True)
        
104         If .MascotaIndex And canWarp Then
106             iMinHP = Npclist(.MascotaIndex).Stats.MinHp
108             PetTiempoDeVida = Npclist(.MascotaIndex).Contadores.TiempoExistencia
110             NpcNumber = Npclist(.MascotaIndex).numero
            
112             Call QuitarNPC(.MascotaIndex)
114             .MascotaIndex = 0
            
                Dim SpawnPos As WorldPos
        
116             SpawnPos.Map = .Pos.Map
118             SpawnPos.X = .Pos.X + RandomNumber(-3, 3)
120             SpawnPos.Y = .Pos.Y + RandomNumber(-3, 3)
        
                'Index = SpawnNpc(NpcNumber, SpawnPos, False, False)
122             Index = CrearNPC(NpcNumber, SpawnPos.Map, SpawnPos)
            
124             If Index = 0 Then
126                 Call WriteConsoleMsg(UserIndex, "Tu mascota no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
                Else
128                 .MascotaIndex = Index

                    ' Nos aseguramos de que conserve el hp, si estaba dañado
130                 Npclist(Index).Stats.MinHp = iMinHP
            
132                 Npclist(Index).MaestroUser = UserIndex
134                 Npclist(Index).Contadores.TiempoExistencia = PetTiempoDeVida
136                 Call FollowAmo(Index)
                End If
            
            End If

        End With
    
138     If Not canWarp Then
140         Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If
    
        '<EhFooter>
        Exit Sub

WarpMascota_Map_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.WarpMascota_Map " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Cerrar_Usuario_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 16/09/2010
        '16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
        '***************************************************
    
100     With UserList(UserIndex)

102         If .flags.UserLogged And Not .Counters.Saliendo Then
104             .Counters.Saliendo = True
106             .Counters.Salir = IIf(((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.Map).Pk), IntervaloCerrarConexion, 0)

108             Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

Cerrar_Usuario_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.Cerrar_Usuario " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo CancelExit_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/02/08
        '
        '***************************************************
100     If UserList(UserIndex).Counters.Saliendo Then

            ' Is the user still connected?
102         If UserList(UserIndex).flags.UserLogged Then
104             UserList(UserIndex).Counters.Saliendo = False
106             UserList(UserIndex).Counters.Salir = 0
108             Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Else
                'Simply reset
110             UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Pk, IntervaloCerrarConexion, 0)
            End If
        End If

        
            'If Teleports create, cancel
             Call Teleports_Cancel(UserIndex)
        '<EhFooter>
        Exit Sub

CancelExit_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.CancelExit " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, _
                       ByVal UserIndexDestino As Integer, _
                       ByVal NuevoNick As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CambiarNick_Err
        '</EhHeader>

        Dim ViejoNick       As String

        Dim ViejoCharBackup As String
    
100     If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
102     ViejoNick = UserList(UserIndexDestino).Name
    
104     If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
            'hace un backup del char
106         ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
108         Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
        End If

        '<EhFooter>
        Exit Sub

CambiarNick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.CambiarNick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserStatsTxtOFF_Err
        '</EhHeader>

        Dim Ups As Single
        Dim Elv As Long, MaxHp As Long, Clase As eClass, Raza As eRaza, Constitucion As Byte
    
        Dim Bronce As Byte, Plata As Byte, Oro As Byte, Premium As Byte
    
100     If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
102         Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Else
104         Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        
106         Elv = val(GetVar(CharPath & Nombre & ".chr", "STATS", "ELV"))
108         MaxHp = val(GetVar(CharPath & Nombre & ".chr", "STATS", "MAXHP"))
110         Clase = val(GetVar(CharPath & Nombre & ".chr", "INIT", "CLASE"))
112         Raza = val(GetVar(CharPath & Nombre & ".chr", "INIT", "RAZA"))
114         Constitucion = val(GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT" & eAtributos.Constitucion))
116         Ups = MaxHp - Mod_Balance.getVidaIdeal(Elv, Clase, Constitucion)
         
            'Bronce = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "Bronce"))
            'Plata = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "PLATA"))
            'Oro = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "ORO"))
            'Premium = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "PREMIUM"))
        
118         Call WriteConsoleMsg(sendIndex, "Nivel: " & Elv & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
120         Call WriteConsoleMsg(sendIndex, "Clase-Raza: " & ListaClases(Clase) & " " & ListaRazas(Raza), FontTypeNames.FONTTYPE_INFO)
        
122         Call WriteConsoleMsg(sendIndex, IIf(Bronce > 0, "BRONCE: SI. ", "BRONCE: NO. ") & IIf(Plata > 0, "PLATA: SI. ", "PLATA: NO. ") & IIf(Premium > 0, "PREMIUM: SI. ", "PREMIUM: NO. ") & Oro, FontTypeNames.FONTTYPE_INFO)
        
            'Call WriteConsoleMsg(sendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
124         Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & MaxHp & " Ups: " & Ups & ",  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
126         Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
128         Call WriteConsoleMsg(sendIndex, "Dsp: " & GetVar(CharPath & Nombre & ".chr", "stats", "ELDHIR"), FontTypeNames.FONTTYPE_INFO)
        
            #If ConUpTime Then

                Dim TempSecs As Long

                Dim TempSTR  As String

130             TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
132             TempSTR = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
134             Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempSTR, FontTypeNames.FONTTYPE_INFO)
            #End If
    
            'Call WriteConsoleMsg(sendIndex, "Dados: " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT1") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT2") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT3") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT4") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT5"), FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

SendUserStatsTxtOFF_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SendUserStatsTxtOFF " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserOROTxtFromChar_Err
        '</EhHeader>

        Dim Charfile As String

        Dim Account  As String
    

100     Charfile = CharPath & charName & ".chr"
    
102     If FileExist(Charfile, vbNormal) Then
104         Account = GetVar(Charfile, "INIT", "ACCOUNTNAME")
        
106         Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "GLD") & " Monedas de Oro en el banco.", FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "ELDHIR") & " Monedas de Eldhir en el banco.", FontTypeNames.FONTTYPE_INFO)
    
        Else
112         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

SendUserOROTxtFromChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SendUserOROTxtFromChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo VolverCriminal_Err
        '</EhHeader>

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/02/2010
        'Nacho: Actualiza el tag al cliente
        '21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
        '**************************************************************
100     With UserList(UserIndex)

102         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
104         If MapInfo(.Pos.Map).FreeAttack = True Then Exit Sub
        
106         If .flags.Privilegios And (PlayerType.User) Then
108             .Reputacion.BurguesRep = 0
110             .Reputacion.NobleRep = 0
112             .Reputacion.PlebeRep = 0
114             .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO

116             If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            
118             If .Faction.Status = r_Armada Then
120                 Call mFacciones.Faction_RemoveUser(UserIndex)
                Else
122                 Call Guilds_CheckAlineation(UserIndex, a_Neutral)
                End If
            
124             If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0

            End If

        End With
    
126     Call RefreshCharStatus(UserIndex)
        '<EhFooter>
        Exit Sub

VolverCriminal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.VolverCriminal " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo VolverCiudadano_Err
        '</EhHeader>

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/06/2006
        'Nacho: Actualiza el tag al cliente.
        '**************************************************************
100     With UserList(UserIndex)

102         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
104         .Reputacion.LadronesRep = 0
106         .Reputacion.BandidoRep = 0
108         .Reputacion.AsesinoRep = 0
110         .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO

112         If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
114         Call Guilds_CheckAlineation(UserIndex, a_Neutral)
        End With
    
116     Call RefreshCharStatus(UserIndex)
        '<EhFooter>
        Exit Sub

VolverCiudadano_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.VolverCiudadano " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal Body As Integer) As Boolean
        '<EhHeader>
        On Error GoTo BodyIsBoat_Err
        '</EhHeader>

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 10/07/2008
        'Checks if a given body index is a boat
        '**************************************************************
        'TODO : This should be checked somehow else. This is nasty....
100     If Body = iFragataReal Or Body = iFragataCaos Or Body = iBarcaPk Or Body = iGaleraPk Or Body = iGaleonPk Or Body = iBarcaCiuda Or Body = iGaleraCiuda Or Body = iGaleonCiuda Or Body = iFragataFantasmal Then
102         BodyIsBoat = True
        End If

        '<EhFooter>
        Exit Function

BodyIsBoat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.BodyIsBoat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, _
                        ByVal userCharIndex As Integer, _
                        ByVal Invisible As Boolean, _
                        Optional ByVal Intermitencia As Boolean = False)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SetInvisible_Err
        '</EhHeader>

        Dim sndNick As String

100     With UserList(UserIndex)
102         Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, Invisible, Intermitencia))
    
104         sndNick = .Name

        
106         If Invisible Then
108             sndNick = sndNick & " " & TAG_USER_INVISIBLE
            
            
            Else
            
110             If .GuildIndex > 0 Then
112                 sndNick = sndNick & " <" & GuildsInfo(.GuildIndex).Name & ">"
                End If
            
114             Call WriteUpdateGlobalCounter(UserIndex, 1, 0)
            End If
    
116         Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, UserIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
        End With

        '<EhFooter>
        Exit Sub

SetInvisible_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SetInvisible " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SetConsulatMode(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/06/10
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SetConsulatMode_Err
        '</EhHeader>

        Dim sndNick As String

100     With UserList(UserIndex)
102         sndNick = .Name
    
104         If EsGm(UserIndex) Then
106             If UCase$(sndNick) <> "LION" Then
108                 sndNick = sndNick & " " & TAG_GAME_MASTER
                End If
            End If
                    
110         If .flags.EnConsulta Then
112             sndNick = sndNick & " " & TAG_CONSULT_MODE
            Else

114             If .GuildIndex > 0 Then
116                 sndNick = sndNick & " <" & GuildsInfo(.GuildIndex).Name & ">"
                End If
            End If
    
118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.charindex, sndNick))
        End With

        '<EhFooter>
        Exit Sub

SetConsulatMode_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SetConsulatMode " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function IsArena(ByVal UserIndex As Integer) As Boolean
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 10/11/2009
        'Returns true if the user is in an Arena
        '**************************************************************
        '<EhHeader>
        On Error GoTo IsArena_Err
        '</EhHeader>
100     IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)
        '<EhFooter>
        Exit Function

IsArena_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.IsArena " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub PerdioNpc(ByVal UserIndex As Integer, _
                     Optional ByVal CheckPets As Boolean = True)
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 11/07/2010 (ZaMa)
        'The user loses his owned npc
        '18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
        '11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
        '13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
        '**************************************************************
        '<EhHeader>
        On Error GoTo PerdioNpc_Err
        '</EhHeader>

        Dim PetCounter As Long

        Dim PetIndex   As Integer

        Dim NpcIndex   As Integer
    
100     With UserList(UserIndex)
        
102         NpcIndex = .flags.OwnedNpc

104         If NpcIndex > 0 Then
            
106             If CheckPets Then
108                 If .MascotaIndex Then

                        ' Si esta atacando al npc deja de hacerlo
110                     If Npclist(.MascotaIndex).TargetNPC = NpcIndex Then
112                         Call FollowAmo(.MascotaIndex)
                        End If
                
                    End If
                End If
            
                ' Reset flags
114             Npclist(NpcIndex).Owner = 0
116             .flags.OwnedNpc = 0

            End If

        End With

        '<EhFooter>
        Exit Sub

PerdioNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.PerdioNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ApropioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 27/07/2010 (zaMa)
        'The user owns a new npc
        '18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
        '19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
        '27/07/2010: ZaMa - El sistema no aplica a mapas seguros.
        '**************************************************************
        '<EhHeader>
        On Error GoTo ApropioNpc_Err
        '</EhHeader>

100     With UserList(UserIndex)

            ' Los admins no se pueden apropiar de npcs
102         If EsGm(UserIndex) Then Exit Sub
        
            Dim mapa As Integer

104         mapa = .Pos.Map
        
            ' No aplica a triggers seguras
106         If MapData(mapa, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        
            ' No se aplica a mapas seguros
108         If MapInfo(mapa).Pk = False Then Exit Sub
        
            ' No aplica a algunos mapas que permiten el robo de npcs
110         If MapInfo(mapa).RoboNpcsPermitido = 1 Then Exit Sub
        
            ' Pierde el npc anterior
112         If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
            ' Si tenia otro dueño, lo perdio aca
114         Npclist(NpcIndex).Owner = UserIndex
116         .flags.OwnedNpc = NpcIndex
        End With
    
        ' Inicializo o actualizo el timer de pertenencia
118     Call IntervaloPerdioNpc(UserIndex, True)
        '<EhFooter>
        Exit Sub

ApropioNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.ApropioNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function GetDireccion(ByVal UserIndex As Integer, _
                             ByVal OtherUserIndex As Integer) As String
        '<EhHeader>
        On Error GoTo GetDireccion_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 17/11/2009
        'Devuelve la direccion hacia donde esta el usuario
        '**************************************************************
        Dim X As Integer

        Dim Y As Integer
    
100     X = UserList(UserIndex).Pos.X - UserList(OtherUserIndex).Pos.X
102     Y = UserList(UserIndex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    
104     If X = 0 And Y > 0 Then
106         GetDireccion = "Sur"
108     ElseIf X = 0 And Y < 0 Then
110         GetDireccion = "Norte"
112     ElseIf X > 0 And Y = 0 Then
114         GetDireccion = "Este"
116     ElseIf X < 0 And Y = 0 Then
118         GetDireccion = "Oeste"
120     ElseIf X > 0 And Y < 0 Then
122         GetDireccion = "NorEste"
124     ElseIf X < 0 And Y < 0 Then
126         GetDireccion = "NorOeste"
128     ElseIf X > 0 And Y > 0 Then
130         GetDireccion = "SurEste"
132     ElseIf X < 0 And Y > 0 Then
134         GetDireccion = "SurOeste"
        End If

        '<EhFooter>
        Exit Function

GetDireccion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.GetDireccion " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function SameFaccion(ByVal UserIndex As Integer, _
                            ByVal OtherUserIndex As Integer) As Boolean
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 17/11/2009
        'Devuelve True si son de la misma faccion
        '**************************************************************
        '<EhHeader>
        On Error GoTo SameFaccion_Err
        '</EhHeader>
100     SameFaccion = (esCaos(UserIndex) And esCaos(OtherUserIndex)) Or (esArmada(UserIndex) And esArmada(OtherUserIndex))
        '<EhFooter>
        Exit Function

SameFaccion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.SameFaccion " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal UserIndex As Integer, _
                         ByVal Skill As Byte, _
                         ByVal Allocation As Boolean)
        '*************************************************
        'Author: Torres Patricio (Pato)
        'Last modified: 11/20/2009
        '
        '*************************************************
        '<EhHeader>
        On Error GoTo CheckEluSkill_Err
        '</EhHeader>

100     With UserList(UserIndex).Stats

102         If .UserSkills(Skill) < MAXSKILLPOINTS Then
104             If Allocation Then
106                 .ExpSkills(Skill) = 0
                Else
108                 .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)
                End If
        
110             .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
            Else
112             .ExpSkills(Skill) = 0
114             .EluSkills(Skill) = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

CheckEluSkill_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.CheckEluSkill " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, _
                               ByVal ObjIndex As Integer, _
                               ByVal Amount As Long) As Boolean
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 25/11/2009
        'Cheks Wether the user has the required amount of items in the inventory or not
        '**************************************************************
        '<EhHeader>
        On Error GoTo HasEnoughItems_Err
        '</EhHeader>

        Dim Slot          As Long

        Dim ItemInvAmount As Long
    
100     With UserList(UserIndex)

102         For Slot = 1 To .CurrentInventorySlots

                ' Si es el item que busco
104             If .Invent.Object(Slot).ObjIndex = ObjIndex Then
                    ' Lo sumo a la cantidad total
106                 ItemInvAmount = ItemInvAmount + .Invent.Object(Slot).Amount
                End If

108         Next Slot

        End With
    
110     HasEnoughItems = Amount <= ItemInvAmount
        '<EhFooter>
        Exit Function

HasEnoughItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.HasEnoughItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, _
                                ByVal UserIndex As Integer) As Long
        '<EhHeader>
        On Error GoTo TotalOfferItems_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 25/11/2009
        'Cheks the amount of items the user has in offerSlots.
        '**************************************************************
        Dim Slot As Byte
    
100     For Slot = 1 To MAX_OFFER_SLOTS

            ' Si es el item que busco
102         If UserList(UserIndex).ComUsu.Objeto(Slot) = ObjIndex Then
                ' Lo sumo a la cantidad total
104             TotalOfferItems = TotalOfferItems + UserList(UserIndex).ComUsu.cant(Slot)
            End If

106     Next Slot

        '<EhFooter>
        Exit Function

TotalOfferItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.TotalOfferItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo getMaxInventorySlots_Err
        '</EhHeader>

100     If UserList(UserIndex).Invent.MochilaEqpObjIndex > 0 Then
102         getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(UserIndex).Invent.MochilaEqpObjIndex).MochilaType * 5 '5=slots por fila, hacer constante
        Else
104         getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End If

        '<EhFooter>
        Exit Function

getMaxInventorySlots_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.getMaxInventorySlots " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


Public Sub goHome(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/2010
        '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo (lo saque de tPiquetec)
        '***************************************************
        '<EhHeader>
        On Error GoTo goHome_Err
        '</EhHeader>

        Dim Distance As Long

        Dim Tiempo   As Long
    
100     With UserList(UserIndex)
        
102         Select Case .Account.Premium
                Case 0
                    Tiempo = 120
                Case 1
104                 Tiempo = 60
110             Case 2
112                 Tiempo = 30
114             Case 3
116                 Tiempo = 5
            End Select
            
        
118         .Counters.goHomeSec = Tiempo
120         Call IntervaloGoHome(UserIndex, Tiempo * 1000, True)
                
           ' If .flags.Navegando = 1 Then
              '  .Char.FX = AnimHogarNavegando(.Char.Heading)
           ' Else
              '  .Char.FX = AnimHogar(.Char.Heading)

         '   End If
                
           ' .Char.loops = INFINITE_LOOPS
           ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
                
122         Call WriteMultiMessage(UserIndex, eMessages.Home, Distance, Tiempo, , MapInfo(Ciudades(.Hogar).Map).Name)
        
        End With
    
        '<EhFooter>
        Exit Sub

goHome_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.goHome " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Sub setHome(ByVal UserIndex As Integer, _
                   ByVal newHome As eCiudad, _
                   ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo setHome_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/2010
        '30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
        '01/06/2010: ZaMa - Ahora te avisa si ya tenes ese hogar.
        '***************************************************
100     If newHome < eCiudad.cUllathorpe Or newHome > eCiudad.cLastCity - 1 Then Exit Sub
          If newHome = eCiudad.cEsperanza And UserList(UserIndex).Stats.Elv >= 35 Then Exit Sub
          
102     If UserList(UserIndex).Hogar <> newHome Then
104         UserList(UserIndex).Hogar = newHome
    
106         Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a nuestra comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.charindex, vbWhite)
        Else
108         Call WriteChatOverHead(UserIndex, "¡¡¡Ya eres miembro de nuestra comunidad!!!", Npclist(NpcIndex).Char.charindex, vbWhite)
        End If

        '<EhFooter>
        Exit Sub

setHome_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.setHome " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function GetHomeArrivalTime(ByVal UserIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo GetHomeArrivalTime_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Calculates the time left to arrive home.
        '**************************************************************
        Dim TActual As Long
    
100     TActual = GetTime
    
102     With UserList(UserIndex)
104         GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001
        End With

        '<EhFooter>
        Exit Function

GetHomeArrivalTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.GetHomeArrivalTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub HomeArrival(ByVal UserIndex As Integer)
        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Teleports user to its home.
        '**************************************************************
        '<EhHeader>
        On Error GoTo HomeArrival_Err
        '</EhHeader>
    
        Dim tX   As Integer

        Dim tY   As Integer

        Dim tMap As Integer
        
        Dim A As Long
        
100     With UserList(UserIndex)

            'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
102         If .flags.Navegando = 1 Then
104             .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
108             .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
            
110             .Char.ShieldAnim = NingunEscudo
112             .Char.WeaponAnim = NingunArma
114             .Char.CascoAnim = NingunCasco

    
                  For A = 1 To MAX_AURAS
116                 .Char.AuraIndex(A) = NingunAura
                  Next A
            
118             .flags.Navegando = 0
            
120             Call WriteNavigateToggle(UserIndex)
                'Le sacamos el navegando, pero no le mostramos a los demás porque va a ser sumoneado hasta ulla.
            End If
        
122         tX = Ciudades(.Hogar).X
124         tY = Ciudades(.Hogar).Y
126         tMap = Ciudades(.Hogar).Map
        
128         Call FindLegalPos(UserIndex, tMap, tX, tY)
130         Call WarpUserChar(UserIndex, tMap, tX, tY, True)
        
132         Call WriteMultiMessage(UserIndex, eMessages.FinishHome)
        
134         Call EndTravel(UserIndex, False)
        
        End With
    
        '<EhFooter>
        Exit Sub

HomeArrival_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.HomeArrival " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EndTravel(ByVal UserIndex As Integer, ByVal Cancelado As Boolean)
        '<EhHeader>
        On Error GoTo EndTravel_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 11/06/2011
        'Ends travel.
        '**************************************************************
100     With UserList(UserIndex)
102         .Counters.goHome = 0
104         .Counters.goHomeSec = 0
106         .flags.Traveling = 0

108         If Cancelado Then Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
110         .Char.FX = 0
112         .Char.loops = 0
        
114         Call WriteUpdateGlobalCounter(UserIndex, 4, 0)
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 0, 0))
        End With

        '<EhFooter>
        
        Exit Sub

EndTravel_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.EndTravel " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function CharIsValid_Invisibilidad(ByVal UserIndex As Integer, _
                                          ByVal sndIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo CharIsValid_Invisibilidad_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
102         CharIsValid_Invisibilidad = (UserIndex = sndIndex)
        
104         If .Faction.Status <> r_None Then
106             CharIsValid_Invisibilidad = (.Faction.Status = UserList(sndIndex).Faction.Status)
            End If
        
108         If .GuildIndex > 0 Then
110             CharIsValid_Invisibilidad = (.GuildIndex = UserList(sndIndex).GuildIndex)
            End If
       
        End With
    
        '<EhFooter>
        Exit Function

CharIsValid_Invisibilidad_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.CharIsValid_Invisibilidad " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function CharIs_Admin(ByVal Name As String) As Boolean
        '<EhHeader>
        On Error GoTo CharIs_Admin_Err
        '</EhHeader>
    
100     Select Case Name
    
            Case "LION": CharIs_Admin = True

102         Case "MELKOR": CharIs_Admin = True
            
            Case "ARAGON": CharIs_Admin = True
            
104         Case Else: CharIs_Admin = False
        
        End Select

        '<EhFooter>
        Exit Function

CharIs_Admin_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.UsUaRiOs.CharIs_Admin " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function ActualizarVelocidadDeUsuario(ByVal UserIndex As Integer, _
                                             ByVal ShiftRunner As Boolean) As Single

    On Error GoTo 0
    
    Dim velocidad As Single, modificadorItem As Single, modificadorHechizo As Single
   
        velocidad = VelocidadNormal

    modificadorItem = 1
    modificadorHechizo = 1
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            
            'velocidad = VelocidadMuerto
            GoTo UpdateSpeed ' Los muertos no tienen modificadores de velocidad

        End If
        
        ' El traje para nadar es considerado barco, de subtipo = 0
        'If (.flags.Navegando > 0) And (.Invent.BarcoObjIndex > 0) Then
        'modificadorItem = ObjData(.Invent.BarcoObjIndex).velocidad
        'End If
        
        ' If (.flags.Montado = 1) And (.Invent.MonturaObjIndex > 0) Then
        'modificadorItem = ObjData(.Invent.MonturaObjIndex).velocidad
        'End If
        
        ' Algun hechizo le afecto la velocidad
        'If .flags.VelocidadHechizada > 0 Then
        '  modificadorHechizo = .flags.VelocidadHechizada
        'End If
        
        velocidad = VelocidadNormal * modificadorItem * modificadorHechizo
UpdateSpeed:
        .Char.speeding = velocidad
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.charindex, .Char.speeding))
        Call WriteVelocidadToggle(UserIndex)
     
    End With

    Exit Function
    
ActualizarVelocidadDeUsuario_Err:

End Function


' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As eHeading)

        On Error GoTo Handler

100     With UserList(UserIndex)

            ' Elegimos un lado al azar
            Dim r As Integer
102         r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

            ' Roto el heading original hacia ese lado
104         Heading = Rotate_Heading(Heading, r)

            ' Intento moverlo para ese lado
106         If MoveUserChar(UserIndex, Heading) Then
                ' Le aviso al usuario que fue movido
108             Call WriteForceCharMove(UserIndex, Heading)
                Exit Sub
            End If
        
            ' Si falló, intento moverlo para el lado opuesto
110         Heading = InvertHeading(Heading)
112         If MoveUserChar(UserIndex, Heading) Then
                ' Le aviso al usuario que fue movido
114             Call WriteForceCharMove(UserIndex, Heading)
                Exit Sub
            End If
        
            ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
            Dim NuevaPos As WorldPos
116         Call ClosestLegalPos(.Pos, NuevaPos, .flags.Navegando = 1, .flags.Navegando = 0)
118         Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, False)

        End With

        Exit Sub
    
Handler:

End Sub
