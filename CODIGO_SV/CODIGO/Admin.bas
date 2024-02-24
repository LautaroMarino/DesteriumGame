Attribute VB_Name = "Admin"
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

' TEST
Public Type tMotd

    Texto As String
    Formato As String

End Type

Public MaxLines As Integer

Public MOTD()   As tMotd

Public Type tAPuestas

    Ganancias As Long
    Perdidas As Long
    Jugadas As Long

End Type

Public Apuestas                          As tAPuestas

Public tInicioServer                     As Long

Public EstadisticasWeb                   As clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar         As Integer

Public StaminaIntervaloSinDescansar      As Integer

Public SanaIntervaloDescansar            As Integer

Public StaminaIntervaloDescansar         As Integer

Public IntervaloSed                      As Integer

Public IntervaloHambre                   As Integer

Public IntervaloVeneno                   As Integer

Public IntervaloParalizado               As Integer

Public Const IntervaloParalizadoReducido As Integer = 37

Public IntervaloInvisible                As Integer

Public IntervaloFrio                     As Integer

Public IntervaloWavFx                    As Integer

Public IntervaloLanzaHechizo             As Integer

Public IntervaloNPCPuedeAtacar           As Integer

Public IntervaloNPCAI                    As Integer

Public IntervaloInvocacion               As Integer

Public IntervaloOculto                   As Integer '[Nacho]

Public IntervaloUserPuedeAtacar          As Long

Public IntervaloGolpeUsar                As Long

Public IntervaloMagiaGolpe               As Long

Public IntervaloGolpeMagia               As Long

Public IntervaloUserPuedeCastear         As Long

Public IntervaloUserPuedeShiftear         As Long

Public IntervaloUserPuedeTrabajar        As Long

Public IntervaloParaConexion             As Long

Public IntervaloCerrarConexion           As Long '[Gonzalo]

Public IntervaloUserPuedeUsar            As Long

Public IntervaloUserPuedeUsarClick       As Long

Public IntervaloFlechasCazadores         As Long

Public IntervaloPuedeSerAtacado          As Long

Public IntervaloAtacable                 As Long

Public IntervaloOwnedNpc                 As Long

Public IntervalDrop                      As Long

Public MaximoSpeedHack As Long
Public IntervaloCaminar                      As Long

Public IntervaloMeditar                      As Long

Public IntervaloPuedeCastear             As Long

Public IntervalCommerce                  As Long

Public IntervalMessage                   As Long

Public IntervalInfoMao                   As Long

Public IntervaloEquipped               As Long


'BALANCE

Public MinutosWs                         As Long

Public IntervaloGuardarUsuarios          As Long

Public IntervaloTimerGuardarUsuarios     As Long

Public Puerto                            As Integer

Public DateAperture As String
Public BootDelBackUp                     As Byte

Public TOLERANCE_MS_POTION               As Integer

Public TOLERANCE_AMOUNT_POTION           As Byte

Public TOLERANCE_POTIONBLUE_CLIC         As Byte

Public TOLERANCE_POTIONBLUE_U            As Byte

Public TOLERANCE_POTIONRED_CLIC          As Byte

Public TOLERANCE_POTIONRED_U             As Byte

Public DeNoche                           As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    VersionOK = (Ver = ULTIMAVERSION)

End Function

Sub ReSpawnOrigPosNpcs()

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ReSpawnOrigPosNpcs_Err

        '</EhHeader>

        Dim i     As Integer

        Dim MiNPC As Npc
       
100     For i = 1 To LastNPC

            'OJO
102         If Npclist(i).flags.NPCActive Then
            
104             If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).numero = Guardias Then
106                 MiNPC = Npclist(i)
108                 Call QuitarNPC(i)
110                 Call RespawnNpc(MiNPC)

                End If
            
                'tildada por sugerencia de yind
                'If Npclist(i).Contadores.TiempoExistencia > 0 Then
                '        Call MuereNpc(i, 0)
                'End If
            End If
       
112     Next i
    
        '<EhFooter>
        Exit Sub

ReSpawnOrigPosNpcs_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Admin.ReSpawnOrigPosNpcs " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub WorldSave()

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo WorldSave_Err

        '</EhHeader>

        Dim LoopX As Integer

        Dim hFile As Integer
    
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))

        Call GuardarUsuarios(True)
        
        ' Guardamos todos los CLANES
110    Call Guilds_Save_All
    
        ' Respawn de los guardias a las posiciones originales
        'Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

112     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
        '<EhFooter>
        Exit Sub

WorldSave_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Admin.WorldSave " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub PurgarPenas()

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo PurgarPenas_Err

        '</EhHeader>

        Dim i As Long
    
100     For i = 1 To LastUser

102         If UserList(i).flags.UserLogged Then
104             If UserList(i).Counters.Pena > 0 Then
106                 UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
108                 If UserList(i).Counters.Pena < 1 Then
110                     UserList(i).Counters.Pena = 0
112                     Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
114                     Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
116                     Call SendData(SendTarget.ToAll, i, PrepareMessageConsoleMsg("El personaje " & UserList(i).Name & " ha salido de la carcel. ¡Esperamos que haya aprendido la lección!", FontTypeNames.FONTTYPE_INFORED))
                    
118                     Call FlushBuffer(i)

                    End If

                End If

            End If

120     Next i

        '<EhFooter>
        Exit Sub

PurgarPenas_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Admin.PurgarPenas " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, _
                      ByVal Minutos As Long, _
                      Optional ByVal GmName As String = vbNullString)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo Encarcelar_Err

        '</EhHeader>

        ' // NUEVO
100     With UserList(UserIndex)
102         AbandonateEvent UserIndex, , True
              
104         If .flags.SlotReto > 0 Then
106             Call mRetos.UserdieFight(UserIndex, 0, True)

            End If
        
108         If .flags.Desafiando > 0 Then
110             Desafio_UserKill UserIndex

            End If
        
112         If .flags.SlotFast > 0 Then
114             RetoFast_UserDie UserIndex, True

            End If
        
116         If .flags.Transform Then
118             Call Transform_User(UserIndex, 0)

            End If
        
120         .Counters.Pena = Minutos
    
122         Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
    
124         If LenB(GmName) = 0 Then
126             Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
            Else
128             Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)

            End If
            
130         If .flags.Traveling = 1 Then
132             Call EndTravel(UserIndex, True)

            End If
            
            
            WriteMessageDiscord CHANNEL_PENAS, "El personaje **" & .Name & "** ha sido encarcelado durante **" & Minutos & " minutos.**"
        
        End With

        '<EhFooter>
        Exit Sub

Encarcelar_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Admin.Encarcelar " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub BorrarUsuario(ByVal UserName As String)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo BorrarUsuario_Err

        '</EhHeader>

100     If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
102         Kill CharPath & UCase$(UserName) & ".chr"

        End If

        '<EhFooter>
        Exit Sub

BorrarUsuario_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Admin.BorrarUsuario " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    BANCheck = (val(GetVar(CharPath & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo PersonajeExiste_Err
        '</EhHeader>
100     PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

        '<EhFooter>
        Exit Function

PersonajeExiste_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.PersonajeExiste " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function UnBan(ByVal Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UnBan_Err
        '</EhHeader>

        'Unban the character
100     Call WriteVar(CharPath & Name & ".chr", "FLAGS", "Ban", "0")
    
        'Remove it from the banned people database
102     Call WriteVar(LogPath & "BanDetail.dat", Name, "BannedBy", "NOBODY")
104     Call WriteVar(LogPath & "BanDetail.dat", Name, "Reason", "NO REASON")

        '<EhFooter>
        Exit Function

UnBan_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.UnBan " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub LoadElu()
        '***************************************************
        'Author: WAICON
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LoadElu_Err
        '</EhHeader>

        Dim A As Integer

100     For A = 1 To STAT_MAXELV
102         EluUser(A) = val(GetVar(IniPath & "Server.ini", "EXPERIENCIA", A))
            'Texto = "^   " & A & "   |   " & EluUser(A) & "         |"
            ' LogError(Texto)
104     Next A

        '<EhFooter>
        Exit Sub

LoadElu_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.LoadElu " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub BanIpAgrega(ByVal IP As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    BanIps.Add IP
    
    Call BanIpGuardar

End Sub

Public Function BanIpBuscar(ByVal IP As String) As Long
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo BanIpBuscar_Err
        '</EhHeader>

        Dim Dale  As Boolean

        Dim LoopC As Long
    
100     Dale = True
102     LoopC = 1

104     Do While LoopC <= BanIps.Count And Dale
106         Dale = (BanIps.Item(LoopC) <> IP)
108         LoopC = LoopC + 1
        Loop
    
110     If Dale Then
112         BanIpBuscar = 0
        Else
114         BanIpBuscar = LoopC - 1

        End If

        '<EhFooter>
        Exit Function

BanIpBuscar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.BanIpBuscar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function BanIpQuita(ByVal IP As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo BanIpQuita_Err
        '</EhHeader>


        Dim N As Long
    
100     N = BanIpBuscar(IP)

102     If N > 0 Then
104         BanIps.Remove N
106         BanIpGuardar
108         BanIpQuita = True
        Else
110         BanIpQuita = False

        End If

        '<EhFooter>
        Exit Function

BanIpQuita_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.BanIpQuita " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub BanIpGuardar()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo BanIpGuardar_Err
        '</EhHeader>

        Dim ArchivoBanIp As String

        Dim ArchN        As Long

        Dim LoopC        As Long
    
100     ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
102     ArchN = FreeFile()
104     Open ArchivoBanIp For Output As #ArchN
    
106     For LoopC = 1 To BanIps.Count
108         Print #ArchN, BanIps.Item(LoopC)
110     Next LoopC
    
112     Close #ArchN

        '<EhFooter>
        Exit Sub

BanIpGuardar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.BanIpGuardar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub BanIpCargar()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo BanIpCargar_Err
        '</EhHeader>

        Dim ArchN        As Long

        Dim Tmp          As String

        Dim ArchivoBanIp As String
    
100     ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
102     Set BanIps = New Collection
    
104     ArchN = FreeFile()
106     Open ArchivoBanIp For Input As #ArchN
    
108     Do While Not EOF(ArchN)
110         Line Input #ArchN, Tmp
112         BanIps.Add Tmp
        Loop
    
114     Close #ArchN

        '<EhFooter>
        Exit Sub

BanIpCargar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.BanIpCargar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
        '***************************************************
        'Author: Unknown
        'Last Modification: 03/02/07
        'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        '***************************************************
        '<EhHeader>
        On Error GoTo UserDarPrivilegioLevel_Err
        '</EhHeader>

100     If EsAdmin(Name) Then
102         UserDarPrivilegioLevel = PlayerType.Admin
104     ElseIf EsDios(Name) Then
106         UserDarPrivilegioLevel = PlayerType.Dios
108     ElseIf EsSemiDios(Name) Then
110         UserDarPrivilegioLevel = PlayerType.SemiDios
        Else
112         UserDarPrivilegioLevel = PlayerType.User

        End If

        '<EhFooter>
        Exit Function

UserDarPrivilegioLevel_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.UserDarPrivilegioLevel " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal Reason As String, _
                        Optional ByVal DataDay As String = vbNullString)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 03/02/07
        '22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
        '***************************************************
        '<EhHeader>
        On Error GoTo BanCharacter_Err
        '</EhHeader>

        Dim tUser     As Integer

        Dim UserPriv  As Byte

        Dim cantPenas As Byte

        Dim Rank      As Integer
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios
    
        Dim Users(0) As String
                        
108     With UserList(bannerUserIndex)

110         If tUser <= 0 Then
112             Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
            
114             If FileExist(CharPath & UserName & ".chr", vbNormal) Then
116                 UserPriv = UserDarPrivilegioLevel(UserName)
                
118                 If (UserPriv And Rank) > (.flags.Privilegios And Rank) Then
120                     Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                    Else

122                     If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
124                         Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                        Else
126                         Call LogBanFromName(UserName, bannerUserIndex, Reason)
128                         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                        
                            'ponemos el flag de ban a 1
130                         Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                            'ponemos la pena
132                         cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
134                         Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
136                         Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, "BANEADO POR " & LCase$(Reason) & " " & Date & " " & Time & IIf(DataDay <> vbNullString, "Hasta fecha: " & DataDay, vbNullString))
138                         Call WriteVar(CharPath & UserName & ".chr", "PENAS", "DATADAY", DataDay)
                        
140                         If (UserPriv And Rank) = (.flags.Privilegios And Rank) Then
142                             .flags.Ban = 1
144                             Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                                'Call CloseSocket(bannerUserIndex)
146                             Call WriteDisconnect(bannerUserIndex)
148                             Call FlushBuffer(bannerUserIndex)
                                                
150                             Call CloseSocket(bannerUserIndex)

                            End If
                        
152                         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "BAN a " & UserName)
                       
                           Dim Account As String
154                        Account = GetVar(CharPath & UserName & ".chr", "INIT", "ACCOUNTNAME")
156                        Call mMao.Mercader_SearchPublications_User(Account, UCase$(UserName))

                        End If

                    End If

                Else
158                 Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

160           If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
162                 Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Else
            
164                 Call LogBan(tUser, bannerUserIndex, Reason)
166                 Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                
                    'Ponemos el flag de ban a 1
168                 UserList(tUser).flags.Ban = 1
                
170                 If (UserList(tUser).flags.Privilegios And Rank) = (.flags.Privilegios And Rank) Then
172                     .flags.Ban = 1
174                     Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                        'Call CloseSocket(bannerUserIndex)
176                     Call WriteDisconnect(bannerUserIndex)
178                     Call FlushBuffer(bannerUserIndex)
                                                
180                     Call CloseSocket(bannerUserIndex)

                    End If
                
182                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "BAN a " & UserName)
                
                    'ponemos el flag de ban a 1
184                 Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
186                 cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
188                 Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
190                 Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & Time & IIf(DataDay <> vbNullString, "Hasta fecha: " & DataDay, vbNullString))
192                 Call WriteVar(CharPath & UserName & ".chr", "PENAS", "DATADAY", DataDay)
                    'Call CloseSocket(tUser)
194                 Call WriteDisconnect(tUser)
196                 Call FlushBuffer(tUser)
                                                
198                 Call CloseSocket(tUser)
200                 Call mMao.Mercader_SearchPublications_User(UserList(tUser).Account.Email, UCase$(UserName))

                End If

            End If

        End With

        '<EhFooter>
        Exit Sub

BanCharacter_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Admin.BanCharacter " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub BanCharacter_Account(ByVal bannerUserIndex As Integer, _
                                ByVal UserName As String, _
                                ByVal Reason As String, _
                                Optional ByVal DataDay As String = vbNullString)

        '<EhHeader>
        On Error GoTo BanCharacter_Account_Err

        '</EhHeader>

        Dim tUser       As Integer

        Dim UserPriv    As Byte

        Dim cantPenas   As Byte
    
        Dim Account     As String

        Dim AccountName As String
        
        Dim FilePath    As String
    
        Dim A           As Long, Chars(1 To ACCOUNT_MAX_CHARS) As String
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104
        AccountName = GetVar(CharPath & UserName & ".chr", "INIT", "ACCOUNTNAME")
        tUser = CheckEmailLogged(AccountName)
106     FilePath = AccountPath & AccountName & ACCOUNT_FORMAT
    
108     With UserList(bannerUserIndex)

116         If val(GetVar(FilePath, "INIT", "BAN")) <> 0 Then
118             Call WriteConsoleMsg(bannerUserIndex, "La cuenta ya se encuentra baneada.", FontTypeNames.FONTTYPE_INFO)
            Else
120             Call LogBanFromName(UserName, bannerUserIndex, Reason)
122             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado la cuenta de " & UserName, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                        
124             Call WriteVar(FilePath, "INIT", "BAN", "1")
                        
                'ponemos la pena
126             cantPenas = val(GetVar(FilePath, "PENAS", "Cant"))
128             Call WriteVar(FilePath, "PENAS", "Cant", cantPenas + 1)
130             Call WriteVar(FilePath, "PENAS", "P" & cantPenas + 1, "BANEADO POR " & LCase$(Reason) & " " & Date & " " & Time & IIf(DataDay <> vbNullString, "Hasta fecha: " & DataDay, vbNullString))
132             Call WriteVar(FilePath, "PENAS", "DATADAY", DataDay)
    
134             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "BAN a la cuenta de " & UserName)
                        
136             Call mMao.Mercader_SearchPublications_User(GetVar(FilePath, "INIT", "ACCOUNTNAME"), vbNullString, True)
    
            End If
            
158         For A = 1 To ACCOUNT_MAX_CHARS
160             Chars(A) = UCase$(.Account.Chars(A).Name)
162         Next A

164         Call mMao.Mercader_SearchPublications_User(AccountName, vbNullString, True)
                    
            If tUser > 0 Then
                Call WriteDisconnect(tUser, True)
166             Call Protocol.Kick(tUser)

            End If

        End With

        '<EhFooter>
        Exit Sub

BanCharacter_Account_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Admin.BanCharacter_Account " & "at line " & Erl
        
        '</EhFooter>
End Sub
