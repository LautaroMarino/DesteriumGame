Attribute VB_Name = "mRetos"
Option Explicit

Public Const VALOR_RETO As Long = 20000



Public Sub Retos_Loop()

On Error GoTo ErrHandler

    Dim A As Long
    'Dim Seconds As Long
    
    'Seconds = Seconds + 1
    
    For A = 1 To MAX_RETOS_SIMULTANEOS

        With Retos(A)
            If .Run Then
                ' Intervalo para evitar que un usuario muera y termine todo de la nada
                If .TimeDescanso > 0 Then
                    .TimeDescanso = .TimeDescanso - 1
                    
                    If .TimeDescanso = 0 Then
                        
                        FinishFight A, .TeamDescanso, .ChangeRound
                        If Not .ChangeRound Then ResetDuelo A
                     End If
                End If
                
               
                ' @Cada 10 Segundos
                'If Seconds Mod 10 = 0 Then
                   ' If .config(eRetoConfig.eEfectoZona) Then
                    '    Call Fight_User_Effect(A)
                   ' End If
                    
                  '  Seconds = 0
               ' End If
    
                ' @Control del Tiempo restante
                If .Time > 0 Then
                    .Time = .Time - 1
                        
                    If .Time <= 0 Then
                        Call ResetDuelo(A)
                        .Time = 0
                    End If
                End If
            End If
        End With
    
    Next A
    
    'If Seconds > 100 Then Seconds = 0
    Exit Sub
ErrHandler:
End Sub

Private Function Fight_ChangeMap(ByVal Slot As Byte, _
                             ByVal Terreno As Byte) As Integer
    

    On Error GoTo ErrHandler
    
    Dim ArenaSlot As Byte
    
    With Retos(Slot)
        ArenaSlot = Arenas_Free(UBound(.User), 1, Terreno)
        
        If ArenaSlot <> 0 Then
            .Terreno = Terreno
            .Arena = ArenaSlot
            Fight_ChangeMap = .Arena
        End If
    End With
    
    Exit Function
ErrHandler:
End Function


Public Sub Reto_ResetUserTemp(ByRef IUser As User)

    Dim A As Long
    Dim NullRetoTemp As tFight
    
    IUser.RetoTemp = NullRetoTemp

End Sub

Private Sub ResetDueloUser(ByVal UserIndex As Integer, _
                           Optional ByVal Deslogged As Boolean = False)

        '<EhHeader>
        On Error GoTo ResetDueloUser_Err

        '</EhHeader>

        On Error GoTo error

100     With UserList(UserIndex)

102         If .Counters.TimeFight > 0 Then
104             .Counters.TimeFight = 0
                  Call WriteRender_CountDown(UserIndex, .Counters.TimeFight)
106             WriteUserInEvent UserIndex
                  
            End If
                
108         With Retos(.flags.SlotReto)
110             .User(UserList(UserIndex).flags.SlotRetoUser).UserIndex = 0
112             .User(UserList(UserIndex).flags.SlotRetoUser).Team = 0
114             .User(UserList(UserIndex).flags.SlotRetoUser).Accepts = 0
116             .User(UserList(UserIndex).flags.SlotRetoUser).Rounds = 0

            End With
              
118         .flags.SlotReto = 0
120         .flags.SlotRetoUser = 255
              .flags.FightTeam = 0
122         StatsDuelos UserIndex
124         WarpPosAnt UserIndex
        End With
          
        Exit Sub

error:
126     LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : ResetDueloUser() userindex: " & UserIndex
        '<EhFooter>
        Exit Sub

ResetDueloUser_Err:
        LogError Err.description & vbCrLf & "in ResetDueloUser " & "at line " & Erl

        '</EhFooter>
End Sub

Private Sub ResetDuelo(ByVal SlotReto As Byte)

    On Error GoTo error

    Dim LoopC As Integer
    Dim NullReto As tFight

    With Retos(SlotReto)
        Arenas(.Arena).Used = False
        
        For LoopC = LBound(.User()) To UBound(.User())
              
            If .User(LoopC).UserIndex > 0 Then
                ResetDueloUser .User(LoopC).UserIndex
            End If
        Next LoopC
            
    End With
    
    Retos(SlotReto) = NullReto
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : ResetDuelo()"
End Sub


Private Function Fight_FreeSlot(ByVal Zona As eTerreno) As Byte
        '<EhHeader>
        On Error GoTo Fight_FreeSlot_Err
        '</EhHeader>

        Dim A As Integer
          
100     For A = 1 To MAX_RETOS_SIMULTANEOS
102         If Not Retos(A).Run Then
104             Fight_FreeSlot = A
                Exit Function
            End If
106     Next A

        '<EhFooter>
        Exit Function

Fight_FreeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.Fight_FreeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Sub Fight_PrepareUsers(ByVal Slot As Byte)
On Error GoTo ErrHandler

    Dim A As Long
    Dim UserIndex As Integer
    
    With Retos(Slot)
        For A = LBound(.User()) To UBound(.User())
            UserIndex = .User(A).UserIndex
            
            If UserIndex > 0 Then
                UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - .Gld - VALOR_RETO
                WriteUpdateGold UserIndex
                
                UserList(UserIndex).flags.SlotReto = Slot
                UserList(UserIndex).flags.SlotRetoUser = A
             
                With UserList(UserIndex)
                    .PosAnt.Map = .Pos.Map
                    .PosAnt.X = .Pos.X
                    .PosAnt.Y = .Pos.Y
                                  
                End With
                
                If A < ((1 + UBound(.User())) / 2) Then
                    .User(A).Team = 2
                    UserList(UserIndex).flags.FightTeam = 2
                Else
                    .User(A).Team = 1
                    UserList(UserIndex).flags.FightTeam = 1
                End If
                
                 Call Reto_ResetUserTemp(UserList(.User(A).UserIndex))
            End If
        Next A
    
    
    End With
    
    Exit Sub
ErrHandler:
    
End Sub

Private Sub RewardUsers(ByVal SlotReto As Byte, ByVal UserIndex As Integer)

    On Error GoTo error
          
    Dim Obj As Obj
          
    With UserList(UserIndex)
              
        .Stats.Gld = .Stats.Gld + (Retos(SlotReto).Gld * 2)
        Call WriteUpdateGold(UserIndex)
                    
    End With
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : RewardUsers()"
End Sub


Private Function CanContinueFight(ByVal UserIndex As Integer) As Boolean

    On Error GoTo error
          
    ' • Si encontramos un personaje vivo el evento continua.
    Dim LoopC        As Integer

    Dim SlotReto     As Byte

    Dim SlotRetoUser As Byte
          
    SlotReto = UserList(UserIndex).flags.SlotReto
    SlotRetoUser = UserList(UserIndex).flags.SlotRetoUser

    CanContinueFight = False
          
    With Retos(SlotReto)
          
        For LoopC = LBound(.User()) To UBound(.User())

            If .User(LoopC).UserIndex > 0 And .User(LoopC).UserIndex <> UserIndex Then
                If .User(SlotRetoUser).Team = .User(LoopC).Team Then

                    With UserList(.User(LoopC).UserIndex)

                        If .flags.Muerto = 0 Then
                            CanContinueFight = True

                            Exit Function

                        End If

                    End With

                End If
                      
            End If

        Next LoopC
              
    End With

    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CanContinueFight()"
End Function

Private Function AttackerFight(ByVal SlotReto As Byte, ByVal TeamUser As Byte) As Integer

    On Error GoTo error

    ' • Buscamos al AttackerIndex (Caso abandono del evento)
    Dim LoopC As Integer
          
    With Retos(SlotReto)

        For LoopC = LBound(.User()) To UBound(.User())

            If .User(LoopC).UserIndex > 0 Then
                If .User(LoopC).Team > 0 And .User(LoopC).Team <> TeamUser Then
                    AttackerFight = .User(LoopC).UserIndex
                    
                    Exit For

                End If
            End If

        Next LoopC

    End With

    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : AttackerFight()"
End Function

Private Function CanAcceptFight(ByVal UserIndex As Integer, _
                                ByVal UserName As String) As Boolean

    On Error GoTo error
          
    Dim SlotTemp  As Byte

    Dim tUser     As Integer
          
    tUser = NameIndex(UserName)
              
    If tUser <= 0 Then
        ' Personaje offline
        CanAcceptFight = False

        Exit Function

    End If
              
    With UserList(tUser)
        
        If UserList(tUser).RetoTemp.Run = False Then
            ' El personaje no mando ninguna solicitud
            Exit Function

        End If
        
        SlotTemp = SearchFight(UserIndex, tUser)
                  
        If SlotTemp = 255 Then
            CanAcceptFight = False

            ' El personaje no te mando ninguna solicitud
            Exit Function

        End If
                  
        If .RetoTemp.User(SlotTemp).Accepts = 1 Then
            ' El personaje ya aceptó.
            CanAcceptFight = False

            Exit Function

        End If
            
        Dim AcceptCant As Byte: AcceptCant = CheckAccepts(.RetoTemp.User)
        
        ' Es el último a aceptar
        If AcceptCant = 1 Then
            If ValidateFight_Users(UserIndex, .RetoTemp) Then
                
                Dim Slot As Integer: Slot = Fight_FreeSlot(Slot)

                If Slot = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Todos los retos han sido ocupados. Intente más tarde por favor...", FontTypeNames.FONTTYPE_INFORED)

                    Exit Function

                Else
                    CanAcceptFight = True
                    GoFight tUser, Slot
                End If
            End If

        Else
            .RetoTemp.User(SlotTemp).Accepts = 1
            CanAcceptFight = True
        End If
          
    End With
              
    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CanAcceptFight()"
End Function

Private Function ValidateFight_Users(ByVal UserIndex As Integer, _
                                     ByRef RetoData As tFight) As Boolean
                                              
    On Error GoTo error
          
    ' • Validamos al Team seleccionado.
          
    Dim LoopC As Integer

    Dim tUser As Integer

    Dim Name  As String
    
    For LoopC = LBound(RetoData.User()) To UBound(RetoData.User())
        Name = UCase$(RetoData.User(LoopC).Name)
        tUser = NameIndex(Name)
        
        If tUser <= 0 Or EsDios(Name) Or EsAdmin(Name) Or EsSemiDios(Name) Then
            'SendMsjUsers "El personaje " & Users(LoopC) & " está offline.", Users()
            WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " está offline", FontTypeNames.FONTTYPE_INFO

            Exit Function

        End If
                  
        With UserList(tUser)
                    
            If .flags.Muerto = 1 Then
                'SendMsjUsers "El personaje " & Users(LoopC) & " está muerto.", Users()
                WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " está muerto.", FontTypeNames.FONTTYPE_INFO

                Exit Function

            End If

            If .flags.Navegando = 1 Then
                'SendMsjUsers "El personaje " & Users(LoopC) & " está muerto.", Users()
                WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " está navegando por los mares.", FontTypeNames.FONTTYPE_INFO

                Exit Function

            End If

            If MapInfo(.Pos.Map).Pk = True Then

                'WriteConsoleMsg UserIndex, "El personaje " & Users(LoopC) & " no está disponible.", FontTypeNames.FONTTYPE_INFO
                'SendMsjUsers "El personaje " & Users(LoopC) & " no está disponible.", Users()
                Exit Function

            End If
                      
            If (.flags.SlotReto > 0) Or (.flags.SlotEvent > 0) Or (.flags.SlotFast > 0) Or (.flags.Desafiando > 0) Then
                'SendMsjUsers "El personaje " & Users(LoopC) & " está en otro evento.", Users()
                WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " está participando en otro evento.", FontTypeNames.FONTTYPE_INFO

                Exit Function

            End If
                      
            If Not Is_Map_valid(tUser) Then
                WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " está participando en otro evento.", FontTypeNames.FONTTYPE_INFO

                Exit Function

            End If
            
            If .flags.Comerciando Then
                'SendMsjUsers "El personaje " & Users(LoopC) & " está comerciando.", Users()
                WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " no está disponible en este momento.", FontTypeNames.FONTTYPE_INFO

                Exit Function

            End If

            If RetoData.config(eRetoConfig.eItems) = 1 Then
                If .Pos.Map <> 133 Then
                    WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " debe estar en la Sala de Comercio, ya que es un enfrentamiento por el inventario del oponente.", FontTypeNames.FONTTYPE_INFO

                    Exit Function

                End If
                    
            End If
                      
            If .Stats.Gld + VALOR_RETO < RetoData.Gld Then
                'SendMsjUsers "El personaje " & .Name & " no tiene las monedas de oro en su billetera.", Users()
                WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " no tiene suficientes monedas.", FontTypeNames.FONTTYPE_INFO

                Exit Function

            End If
             

            If RetoData.config(eRetoConfig.eCascos) = 0 Then
                If .Invent.CascoEqpSlot > 0 Then
                    WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " tiene equipado el Casco", FontTypeNames.FONTTYPE_INFO

                    Exit Function

                End If

            End If

            If RetoData.config(eRetoConfig.eEscudos) = 0 Then
                If .Invent.EscudoEqpSlot > 0 Then
                    WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " tiene equipado el Escudo.", FontTypeNames.FONTTYPE_INFO

                    Exit Function

                End If

            End If
            
            'If User_TieneObjetos_Especiales(UserIndex, RetoData.config(eRetoConfig.eBronce), RetoData.config(eRetoConfig.ePlata), RetoData.config(eRetoConfig.eOro), RetoData.config(eRetoConfig.ePremium)) Then
                                                
              '  WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " posee un objeto no permitido en el duelo pactado.", FontTypeNames.FONTTYPE_INFO
              '  Exit Function
                    
            'End If
            
            If RetoData.config(eRetoConfig.eItems) > 0 Then
                If .flags.ClainObject > 0 Then
                    WriteConsoleMsg UserIndex, "El personaje " & RetoData.User(LoopC).Name & " todavía no reclamo su premio anterior.", FontTypeNames.FONTTYPE_INFO

                    Exit Function

                End If

            End If
                    
        End With

    Next LoopC
          
    ValidateFight_Users = True
          
    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight_Users()"

End Function

Public Function HayRepetidos(lista() As tFightUser) As Boolean

    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim j As Integer
    
    ' Recorrer la lista
    For i = 0 To UBound(lista) - 1
        ' Verificar si el elemento actual está repetido en la lista
        For j = i + 1 To UBound(lista)
            If StrComp(lista(i).Name, lista(j).Name, vbTextCompare) = 0 Then
                ' Elemento repetido encontrado, retornar False
                HayRepetidos = True
                Exit Function
            End If
        Next j
    Next i
    
    ' No se encontraron elementos repetidos, retornar True
    HayRepetidos = False
    
    Exit Function
ErrHandler:
    
End Function

Private Function ValidateFight_Hack(ByVal UserIndex As Integer, _
                                    ByRef Fight As tFight) As Boolean
        '<EhHeader>
        On Error GoTo ValidateFight_Hack_Err
        '</EhHeader>
    
    
100     With Fight
102         If .Gld < 50000 Or .Gld > 100000000 Then
104             LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos::GLD"
    
                Exit Function
    
            End If

    
106         If .Time <> 300 And .Time <> 600 And .Time <> 1200 And .Time <> 1800 And .Time <> 3600 Then
108             LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos::TIME"
    
                Exit Function
    
            End If
    
110         If .RoundsLimit <> 1 And .RoundsLimit <> 3 And .RoundsLimit <> 5 And .RoundsLimit <> 10 And .RoundsLimit <> 20 Then
112             LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos::ROUNDS"
    
                Exit Function
    
            End If
        
            Dim LongitudUsers As Integer
118         LongitudUsers = UBound(.User) + 1
        
120         If LongitudUsers <> 2 And LongitudUsers <> 4 And LongitudUsers <> 6 And LongitudUsers <> 8 And LongitudUsers <> 10 Then
122             Call WriteConsoleMsg(UserIndex, "Solicitud Inválida. Has puesto mal los nombres de los personajes.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function
            End If
        
124         If UBound(.User) > 1 Then
126             If .Tipo = 4 Then
128                 LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos::PLANTES"
    
                    Exit Function
    
                End If
            
130             If .config(eRetoConfig.eItems) > 0 Then
132                 LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos::ITEMS"
    
                    Exit Function
    
                End If
            End If
        
134         If UBound(.User) <= 1 Then
136             If .config(eRetoConfig.eResucitar) > 0 Then
138                 LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos::RESUCITAR"
    
                    Exit Function
    
                End If
            
            End If
        
140         If UserList(UserIndex).Counters.FightSend > 0 Then
142             Call WriteConsoleMsg(UserIndex, "Has enviado una solicitud recientemente. Aguarda unos instantes...", FontTypeNames.FONTTYPE_INFORED)
    
                Exit Function
    
            End If
        
            Dim A     As Long
    
            Dim tUser As Integer
            Dim FoundSender As Boolean
            Dim User() As tFightUser
            
            If HayRepetidos(.User) Then Exit Function
            
144         ReDim User(LBound(.User) To UBound(.User)) As tFightUser
        
146         For A = LBound(.User) To UBound(.User)
148             tUser = NameIndex(.User(A).Name)
            
150             If StrComp(UCase$(.User(A).Name), UCase$(UserList(UserIndex).Name)) = 0 Then
152                 FoundSender = True
                End If
            
154             If EsDios(.User(A).Name) Or EsAdmin(.User(A).Name) Or EsSemiDios(.User(A).Name) Then
156                 Call WriteConsoleMsg(UserIndex, "Uno de los personajes no está en Fuerte Valhalla.", FontTypeNames.FONTTYPE_INFORED)
    
                    Exit Function
    
                End If
            
158             If tUser <= 0 Then
160                 Call WriteConsoleMsg(UserIndex, "Uno de los personajes no está disponible.", FontTypeNames.FONTTYPE_INFORED)
    
                    Exit Function
    
                End If
            
            
                'If UserList(tUser).Pos.Map <> 88 Then
                    'Call WriteConsoleMsg(UserIndex, "Uno de los personajes no está en Fuerte Valhalla.", FontTypeNames.FONTTYPE_INFORED)
    
                    'Exit Function
    
                'End If
            
162             If tUser > 0 Then
164                 User(A).UserIndex = tUser
166                 User(A).Name = UserList(tUser).Name
                
                    ' El ya aceptó
168                 If tUser = UserIndex Then
170                     User(A).Accepts = 1
                    
                    End If

172                 If UserList(tUser).Counters.FightInvitation > 0 Then
174                     Call WriteConsoleMsg(UserIndex, "El personaje " & .User(A).Name & " está por aceptar otra invitación", FontTypeNames.FONTTYPE_INFORED)
    
                        Exit Function
    
                    End If
                End If
    
176         Next A
        
178         If Not FoundSender Then
180             ValidateFight_Hack = False
                Exit Function
            Else
182             .User = User 'Cargamos la temporal validada
            End If
        
        End With
    
184     ValidateFight_Hack = True
        '<EhFooter>
        Exit Function

ValidateFight_Hack_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.ValidateFight_Hack " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function StrTeam(ByRef FightSlot As Byte) As String
          
    On Error GoTo error
          
    ' • Devuelve ENEMIGOS vs TEAM
          
    Dim LoopC      As Integer

    Dim strTemp(1) As String
    
    With Retos(FightSlot)
        ' 1 vs 1
        If UBound(.User()) = 1 Then
            If .User(0).UserIndex > 0 Then
                strTemp(0) = UserList(.User(0).UserIndex).Name
            Else
                strTemp(0) = "Usuario descalificado"
            End If
                  
            If .User(1).UserIndex > 0 Then
                strTemp(1) = UserList(.User(1).UserIndex).Name
            Else
                strTemp(1) = "Usuario descalificado"
            End If
                  
            StrTeam = strTemp(0) & " vs " & strTemp(1)
    
            Exit Function
    
        End If
    
        For LoopC = LBound(.User()) To UBound(.User())

            If .User(LoopC).UserIndex > 0 Then
                If LoopC < ((1 + UBound(.User)) / 2) Then
                    strTemp(0) = strTemp(0) & UserList(.User(LoopC).UserIndex).Name & ", "
                Else
                    strTemp(1) = strTemp(1) & UserList(.User(LoopC).UserIndex).Name & ", "
                End If
            End If
    
        Next LoopC
        
        If Not strTemp(0) = vbNullString Then
            strTemp(0) = Left$(strTemp(0), Len(strTemp(0)) - 2)
        Else
            strTemp(0) = "Equipo descalificado"
        End If
          
        If Not strTemp(1) = vbNullString Then
            strTemp(1) = Left$(strTemp(1), Len(strTemp(1)) - 2)
        Else
            strTemp(1) = "Equipo descalificado"
        End If
          
        StrTeam = strTemp(0) & " vs " & strTemp(1)
    End With

   

          
    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : StrTeam()"
End Function

Private Function CheckAccepts(ByRef User() As tFightUser) As Byte

    On Error GoTo error
          
    ' • Si encontramos a un usuario que no haya aceptado retornamos false.
    Dim A As Integer

    Dim Temp  As Byte
          
    For A = LBound(User()) To UBound(User())

        If User(A).Accepts = 0 Then
            Temp = Temp + 1
        End If

    Next A
          
    CheckAccepts = Temp

    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CheckAccepts()"
End Function

Private Function SearchFight(ByVal UserIndex As Integer, ByVal tUser As Integer) As Byte
                                      
    ' • Buscamos la invitación que nos realizo el personaje UserName
          
    On Error GoTo error

    Dim A As Integer
          
    SearchFight = 255
          
    With UserList(tUser)
        For A = LBound(.RetoTemp.User()) To UBound(.RetoTemp.User())
            If .RetoTemp.User(A).UserIndex = UserIndex Then
                SearchFight = A
                Exit Function
            End If
        Next A
    End With
    
    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : SearchFight()"
End Function

Public Function CanAttackReto(ByVal AttackerIndex As Integer, _
                              ByVal VictimIndex As Integer) As Boolean
          
    On Error GoTo error

    CanAttackReto = True
          
    With UserList(AttackerIndex)

        If .flags.SlotReto > 0 Then
                  
            'If Retos(.flags.SlotReto).User(.flags.SlotRetoUser).Team = _
             Retos(.flags.SlotReto).User(UserList(VictimIndex).flags.SlotRetoUser).Team Then
            CanAttackReto = True

            Exit Function

            'End If
        End If
          
    End With
          
    Exit Function

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CanAttackReto()"
End Function

Private Sub SendInvitation(ByVal UserIndex As Integer, _
                           ByRef Fight As tFight)
                                  
    On Error GoTo error
          
    ' • Enviamos la solicitud del duelo a los demás y guardamos los datos temporales al usuario mandatario.

    Dim A       As Long
    Dim tUser   As Integer
    
    With UserList(UserIndex)
        .Counters.FightSend = 10
        .RetoTemp = Fight
        .RetoTemp.Run = True
    End With

    Dim TextUsers As String

    TextUsers = Fight_TextUsers(Fight.User)
        
    For A = LBound(Fight.User()) To UBound(Fight.User)
        tUser = Fight.User(A).UserIndex
        
        If tUser > 0 And tUser <> UserIndex Then
            Call WriteFight_PanelAccept(tUser, UserList(UserIndex).Name, TextUsers, UserList(UserIndex).RetoTemp)
            UserList(tUser).Counters.FightInvitation = 5
        End If

    Next A
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : SendInvitation()"
End Sub

Private Function Fight_TextUsers(ByRef User() As tFightUser) As String
        '<EhHeader>
        On Error GoTo Fight_TextUsers_Err
        '</EhHeader>

        Dim Text  As String

        Dim A     As Long

        Dim tUser As Integer
    
100     For A = LBound(User) To UBound(User)
102         tUser = User(A).UserIndex
        
104         If tUser > 0 Then
106             Text = Text & User(A).Name & " (" & ListaClases(UserList(tUser).Clase) & " " & ListaRazas(UserList(tUser).Raza) & ") Lvl " & UserList(tUser).Stats.Elv & vbCrLf & SEPARATOR
            End If

108     Next A
    
110     If LenB(Text) > 0 Then Text = Left$(Text, Len(Text) - 1)
    
112     Fight_TextUsers = Text
        '<EhFooter>
        Exit Function

Fight_TextUsers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.Fight_TextUsers " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


Private Sub GoFight(ByVal UserIndex As Integer, ByRef Slot As Integer)
    ' • Comienzo del duelo
          
    On Error GoTo error
    
    Dim SlotArena As Integer
    
    SlotArena = Fight_ChangeMap(Slot, UserList(UserIndex).RetoTemp.Terreno)
    
    If SlotArena > 0 Then
        With UserList(UserIndex)
            Retos(Slot) = .RetoTemp
            Retos(Slot).Arena = SlotArena
            
            Call Fight_PrepareUsers(Slot)
            
            Call WarpFight(Slot, 1)
        End With
    Else
        Call WriteConsoleMsg(UserIndex, "Has encontrado disponibilidad para disputar un reto, pero el tipo de zona que has elegido se encuentra saturado. Escoge otro.", FontTypeNames.FONTTYPE_INFORETOS)
    End If
    
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : GoFight() " & Retos(Slot).RoundsLimit & "-" & Retos(Slot).Terreno
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : GoFight()"
End Sub

Private Sub WarpFight(ByVal Slot As Byte, ByVal OrdenTeam As Byte)
        '<EhHeader>
        On Error GoTo WarpFight_Err
        '</EhHeader>

        ' • Teletransportamos a los personajes a la sala de combate
          
        On Error GoTo error

        Dim LoopC        As Integer

        Dim tUser        As Integer

        Dim Pos          As WorldPos

        Dim SlotArena    As Integer
    
        Const Tile_Extra As Byte = 5
          
100     With Retos(Slot)
102         For LoopC = LBound(.User()) To UBound(.User())
104             tUser = .User(LoopC).UserIndex
                  
106             If tUser > 0 Then
108                 If Not OrdenTeam = 0 Then ' No warpeamos pero stopeamos
110                     Slot = UserList(tUser).flags.SlotReto
112                     SlotArena = Retos(Slot).Arena
                    
114                     Pos.Map = Arenas(SlotArena).Map
                    
116                     If .User(LoopC).Team = OrdenTeam Then
118                         Pos.X = Arenas(SlotArena).X
120                         Pos.Y = Arenas(SlotArena).Y
                        Else
121                         Pos.X = Arenas(SlotArena).X + Arenas(SlotArena).TileAddX
122                         Pos.Y = Arenas(SlotArena).Y + Arenas(SlotArena).TileAddY
                        End If
                          
126                     With UserList(tUser)
128                         ClosestStablePos Pos, Pos
130                         WarpUserChar tUser, Pos.Map, Pos.X, Pos.Y, False
                        End With
        
                    End If
                
132                 UserList(tUser).Counters.TimeFight = 10
134                 Call WriteUserInEvent(tUser)
                End If
136         Next LoopC
        End With
          
        Exit Sub

error:
138     LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : WarpFight()"
        '<EhFooter>
        Exit Sub

WarpFight_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.WarpFight " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub AddRound(ByVal SlotReto As Byte, ByVal Team As Byte)

    On Error GoTo error

    Dim LoopC As Integer

    With Retos(SlotReto)

        For LoopC = LBound(.User()) To UBound(.User())

            If .User(LoopC).Team = Team And .User(LoopC).UserIndex > 0 Then
                .User(LoopC).Rounds = .User(LoopC).Rounds + 1
            End If

        Next LoopC
          
    End With
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : AddRound()"
End Sub

Private Sub SendMsjUsers(ByVal strMsj As String, ByRef Users() As String)
                              
    On Error GoTo error

    Dim LoopC As Integer

    Dim tUser As Integer
          
    For LoopC = LBound(Users()) To UBound(Users())
        tUser = NameIndex(Users(LoopC))

        If tUser > 0 Then
            WriteConsoleMsg tUser, strMsj, FontTypeNames.FONTTYPE_VENENO
        End If

    Next LoopC
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : SendMsjUsers()"
End Sub

Private Function ExistCompañero(ByVal UserIndex As Integer) As Boolean

    Dim LoopC        As Integer

    Dim SlotReto     As Byte

    Dim SlotRetoUser As Byte
          
    On Error GoTo ExistCompañero_Error

    SlotReto = UserList(UserIndex).flags.SlotReto
    SlotRetoUser = UserList(UserIndex).flags.SlotRetoUser
          
    With Retos(SlotReto)

        For LoopC = LBound(.User()) To UBound(.User())

            If .User(LoopC).UserIndex > 0 Then
                If LoopC <> SlotRetoUser Then
                    If .User(LoopC).Team = .User(SlotRetoUser).Team Then
                        ExistCompañero = True

                        Exit For

                    End If
                End If
            End If

        Next LoopC

    End With

    On Error GoTo 0

    Exit Function

ExistCompañero_Error:

    LogRetos "Error " & Err.number & " (" & Err.description & ") in procedure ExistCompañero of Módulo mRetos in line " & Erl
          
End Function
Public Sub UserdieFight(ByVal UserIndex As Integer, _
                        ByVal AttackerIndex As Integer, _
                        ByVal Forzado As Boolean)

    On Error GoTo error

    ' • Un personaje en reto es matado por otro.
    Dim LoopC     As Integer

    Dim strTemp   As String

    Dim SlotReto  As Byte

    Dim TeamUser  As Byte

    Dim Rounds    As Byte

    Dim Deslogged As Boolean

    Dim ExistTeam As Boolean
          
    SlotReto = UserList(UserIndex).flags.SlotReto
          
    Deslogged = False
10
    ' • Caso hipotetico de deslogeo. El funcionamiento es el mismo, con la diferencia de que se busca al ganador.
    If AttackerIndex = 0 Then
        AttackerIndex = AttackerFight(SlotReto, Retos(SlotReto).User(UserList(UserIndex).flags.SlotRetoUser).Team)
              
        Deslogged = True
    End If
20
    TeamUser = Retos(SlotReto).User(UserList(AttackerIndex).flags.SlotRetoUser).Team
    ExistTeam = ExistCompañero(UserIndex)
30
    ' Deslogeo de todos los integrantes del team
    If Forzado Then
        If Not ExistTeam Then
            If Retos(SlotReto).config(eRetoConfig.eItems) = 1 Then
                Call Retos_SetObj(AttackerIndex, UserIndex)
            End If
40
            FinishFight SlotReto, TeamUser
            ResetDuelo SlotReto
                
            Exit Sub

        End If
    End If
50
    With UserList(UserIndex)

        If Not CanContinueFight(UserIndex) Then

            With Retos(SlotReto)
60
                For LoopC = LBound(.User()) To UBound(.User())
70
                    With .User(LoopC)
80
                        If .UserIndex > 0 And .Team = TeamUser Then
                            If Rounds = 0 Then
                                AddRound SlotReto, .Team
                                Rounds = .Rounds
                            End If
90
                            WriteConsoleMsg .UserIndex, "Has ganado el round. Rounds ganados: " & .Rounds & ".", FontTypeNames.FONTTYPE_VENENO
100
                        End If

                    End With
                          
                    
                Next LoopC
120
                If Rounds >= (.RoundsLimit / 2) + 0.5 Then
                    If Retos(SlotReto).config(eRetoConfig.eItems) = 1 Then
                        Call Retos_SetObj(AttackerIndex, UserIndex)
                    End If
130

                   
                    'FinishFight SlotReto, TeamUser
                     .TeamDescanso = TeamUser
                     .ChangeRound = False
                      .TimeDescanso = 2
                    'ResetDuelo SlotReto

                    Exit Sub

                Else
140
                    .TeamDescanso = TeamUser
                    .ChangeRound = True
                    .TimeDescanso = 2
                    
                    'FinishFight SlotReto, TeamUser, True

                End If
150
            End With

        End If
160

        If Deslogged Then
            ResetDueloUser UserIndex, True
        End If

    End With
165
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : UserdieFight() en linea " & Erl
End Sub

Private Sub StatsDuelos(ByVal UserIndex As Integer)

    On Error GoTo error

    With UserList(UserIndex)

        If .flags.Muerto Then
            RevivirUsuario (UserIndex)
            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMan = .Stats.MaxMan
            .Stats.MinSta = .Stats.MaxSta
              
            WriteUpdateUserStats UserIndex

            Exit Sub

        End If

        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinSta = .Stats.MaxSta
              
        WriteUpdateUserStats UserIndex
            
        'If .flags.Paralizado = 1 Then
        '.flags.Paralizado = 0
        'Call WriteParalizeOK(UserIndex)
        'End If
            
    End With
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : StatsDuelos()"
End Sub

Private Sub FinishFight(ByVal SlotReto As Byte, _
                        ByVal Team As Byte, _
                        Optional ByVal ChangeTeam As Boolean, _
                        Optional ByVal UserDead As Integer = 0)

    ' • Finalizamos el reto o el round.
          
    On Error GoTo error

    Dim LoopC   As Integer

    Dim strTemp As String
          
    With Retos(SlotReto)

        For LoopC = LBound(.User()) To UBound(.User())

            If .User(LoopC).UserIndex > 0 Then
                
                StatsDuelos .User(LoopC).UserIndex
                
                If Team = .User(LoopC).Team Then
                    If ChangeTeam Then
                        StatsDuelos .User(LoopC).UserIndex
                    Else
                        
                        Arenas(.Arena).Used = False
                        StatsDuelos .User(LoopC).UserIndex
                        RewardUsers SlotReto, .User(LoopC).UserIndex
                              
                        If .User(LoopC).Rounds > 0 Then
                            WriteConsoleMsg .User(LoopC).UserIndex, "Has ganado el reto con " & .User(LoopC).Rounds & " rounds a tu favor.", FontTypeNames.FONTTYPE_VENENO
                        Else
                            WriteConsoleMsg .User(LoopC).UserIndex, "Has ganado el reto.", FontTypeNames.FONTTYPE_VENENO
                        End If

                        strTemp = strTemp & UserList(.User(LoopC).UserIndex).Name & ", "

                        If .config(eRetoConfig.eItems) = 1 Then
                            Call WriteConsoleMsg(.User(LoopC).UserIndex, "Has ganado el inventario de tu adversario. Reclama los objetos desde la ciudad principal con comando /RECLAMAR", FontTypeNames.FONTTYPE_INFOGREEN)
                        End If
                        
                    End If
                    
                    
                End If
                
                If .User(LoopC).UserIndex > 0 Then
                    If UBound(.User) = 1 Then
                        Call WriteUpdateEvent(UserList(.User(LoopC).UserIndex).ID, _
                        UserList(.User(LoopC).UserIndex).Name, _
                        eSubType_Modality.Retos1vs1, _
                        Team = .User(LoopC).Team)
                    End If
                End If
                
                
            End If

        Next LoopC
          
        If ChangeTeam Then
            
           ' If .config(eRetoConfig.eCambioZona) = 1 Then
             '   Dim Zona As Integer, MapAnt As Integer
              '  MapAnt = .Arena
             '   Zona = .Zona + 1
                
             '   If Zona = 6 Then Zona = 1
                
             '   If Fight_ChangeMap(SlotReto, Zona) Then
               '     Arena(MapAnt).Run = False
             '   End If
            'End If
            
            Call WarpFight(SlotReto, IIf((.Tipo = 4), 0, 2))
        Else
            
            strTemp = Left$(strTemp, Len(strTemp) - 2)
        
            SendData SendTarget.toMap, 1, PrepareMessageConsoleMsg( _
               IIf((.Tipo = 4), "«Plantes» ", "«Retos» ") & IIf((.RoundsLimit = 1), "¡A 1 Round! ", "¡A " & .RoundsLimit & " Rounds! ") & StrTeam(SlotReto) & ". Ganador " & strTemp & ". Apuesta por " & _
               .Gld & " Monedas de Oro." & IIf((.config(eRetoConfig.eItems) > 0), " ¡¡Le usurpó el inventario!!", vbNullString), FontTypeNames.FONTTYPE_INFORETOS)
                
            LogRetos "«Retos» Rounds ¡" & .RoundsLimit & "! " & StrTeam(SlotReto) & ". Ganador el team de " & strTemp & ". Apuesta por  " & .Gld & " Monedas de Oro." & IIf((.config(eRetoConfig.eItems) > 0), " ¡¡Le usurpó el inventario!", vbNullString)
        End If

    End With
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : FinishFight() en linea " & Erl
End Sub

' El personaje Envia una solicitud de reto

Public Sub SendFight(ByVal UserIndex As Integer, _
                     ByRef Fight As tFight)
          
    On Error GoTo error
          
    With UserList(UserIndex)
        
        If ConfigServer.ModoRetos = 0 Then
            WriteConsoleMsg UserIndex, "El juego no tiene habilitado los enfrentamientos.", FontTypeNames.FONTTYPE_WARNING
            Exit Sub
        End If
        
        If ValidateFight_Hack(UserIndex, Fight) Then
            
            SendInvitation UserIndex, Fight
                  
            WriteConsoleMsg UserIndex, "Has enviado correctamente la solicitud a un duelo. Recuerda que si vuelves a mandar, la anterior solicitud se cancela.", FontTypeNames.FONTTYPE_WARNING
                  
        End If
              
    End With
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : SendFight()"
End Sub

' El personaje acepta una solicitud recibida
Public Sub AcceptFight(ByVal UserIndex As Integer, ByVal UserName As String)
                              
    On Error GoTo error
                              
    With UserList(UserIndex)
        .Counters.FightInvitation = 0
        
        If ConfigServer.ModoRetos = 0 Then
            WriteConsoleMsg UserIndex, "El juego no tiene habilitado los enfrentamientos.", FontTypeNames.FONTTYPE_WARNING
            Exit Sub
        End If
        
        If CanAcceptFight(UserIndex, UserName) Then
            
            WriteConsoleMsg UserIndex, "Has aceptado la invitación.", FontTypeNames.FONTTYPE_INFO

            Dim tUser As Integer: tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, "El personaje " & .Name & " aceptó la invitación al duelo.", FontTypeNames.FONTTYPE_GUILD)
            End If
            
        End If

    End With
          
    Exit Sub

error:
    LogRetos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : AcceptFight()"
End Sub

' Objetos que no pueden ser reclamados por el ganador.
Private Function Retos_Obj_NotWin(ByVal ObjIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Retos_Obj_NotWin_Err
        '</EhHeader>

100     With ObjData(ObjIndex)

102         If ObjData(ObjIndex).OBJType <> otBarcos And _
               ObjData(ObjIndex).OBJType <> otPociones And _
               ObjData(ObjIndex).OBJType <> otBebidas And _
               ObjData(ObjIndex).OBJType <> otUseOnce And _
               ObjData(ObjIndex).OBJType <> otBotellaVacia And _
               ObjData(ObjIndex).OBJType <> otLeña And _
               ObjData(ObjIndex).OBJType <> otMinerales And _
               ObjData(ObjIndex).OBJType <> otBotellaLlena And _
               ObjData(ObjIndex).Real <> 1 And _
               ObjData(ObjIndex).Caos <> 1 And _
               ObjData(ObjIndex).Premium <> 1 And _
               ObjData(ObjIndex).Plata <> 1 And _
               ObjData(ObjIndex).Oro <> 1 And _
               ObjData(ObjIndex).NoNada <> 1 Then
                    
104             Retos_Obj_NotWin = True

                Exit Function

            End If
        
        End With

        '<EhFooter>
        Exit Function

Retos_Obj_NotWin_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.Retos_Obj_NotWin " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub Retos_SetObj(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
        '<EhHeader>
        On Error GoTo Retos_SetObj_Err
        '</EhHeader>
    
        Dim A        As Long

        Dim ObjIndex As Integer

        Dim NuevaPos As WorldPos

        Dim Obj      As Obj
    
100     With UserList(VictimIndex)

102         For A = 1 To .CurrentInventorySlots
104             ObjIndex = .Invent.Object(A).ObjIndex
            
106             If ObjIndex > 0 Then
108                 If Retos_Obj_NotWin(ObjIndex) Then
                    
110                     UserList(UserIndex).ObjectClaim(A).ObjIndex = .Invent.Object(A).ObjIndex
112                     UserList(UserIndex).ObjectClaim(A).Equipped = 0
114                     UserList(UserIndex).ObjectClaim(A).Amount = .Invent.Object(A).Amount

116                     Call QuitarUserInvItem(VictimIndex, A, .Invent.Object(A).Amount)
                    End If
                End If

118         Next A
        
120         Call UpdateUserInv(True, VictimIndex, 0)
        
122         UserList(UserIndex).flags.ClainObject = 1
        
        End With
    
        '<EhFooter>
        Exit Sub

Retos_SetObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.Retos_SetObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub Retos_ReclameObj(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Retos_ReclameObj_Err
        '</EhHeader>
    
        Dim A        As Long

        Dim ObjIndex As Integer

        Dim Obj      As Obj
    
100     With UserList(UserIndex)

102         For A = 1 To MAX_INVENTORY_SLOTS
104             ObjIndex = .ObjectClaim(A).ObjIndex
            
106             If ObjIndex > 0 Then
108                 Obj.ObjIndex = ObjIndex
110                 Obj.Amount = .ObjectClaim(A).Amount
                
112                 If MeterItemEnInventario(UserIndex, Obj) Then
114                     Call WriteConsoleMsg(UserIndex, "Has reclamado: " & ObjData(ObjIndex).Name & " (x" & Obj.Amount & ")", FontTypeNames.FONTTYPE_INFOGREEN)
                    Else
116                     Call Logs_User(.Name, eLog.eUser, eReclameObj, .Name & " no recibió el objeto n°" & ObjIndex & "(" & ObjData(ObjIndex).Name & ") x" & Obj.Amount)
118                     Call LogRetos("El personaje " & .Name & " con IP: " & .IpAddress & " reclamo un objeto el cual no pudo ser entregado por falta de espacio.")
                    End If
                End If
            
120             .ObjectClaim(A).ObjIndex = 0
122             .ObjectClaim(A).Amount = 0
124             .ObjectClaim(A).Equipped = 0
126         Next A
        
128         .flags.ClainObject = 0
        End With

        '<EhFooter>
        Exit Sub

Retos_ReclameObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.Retos_ReclameObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Retos_Reset_All()
        '<EhHeader>
        On Error GoTo Retos_Reset_All_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_RETOS_SIMULTANEOS
102         If Retos(A).Run Then
104             Call ResetDuelo(A)
            End If
106     Next A

        '<EhFooter>
        Exit Sub

Retos_Reset_All_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetos.Retos_Reset_All " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
