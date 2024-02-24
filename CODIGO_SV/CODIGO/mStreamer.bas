Attribute VB_Name = "mStreamer"
Option Explicit

' Modulo encargado de tener un GM BOT que vaya siendo sumoneado de manera automática por los diferentes eventos del juego.

' @ Tiempo entre un Summon y OTRO [VALOR DEFAULT]
Public Const STREAMER_TIME_AUTO_WARP As Long = 10000 ' 6s

' @ Tiempo para que el mismo usuario sea buscado.
Public Const STREAMER_TIME_CAN_SEARCH As Long = 120000 ' 30s

' @ Maximo de usuarios que el stream va a tener en la lista para seguir.
Public Const MAX_STREAM_USERS As Byte = 50

Public Enum eStreamerMode
    eZonaSegura = 1         ' Busca personajes en zona insegura
    eEventos = 2               ' Eventos automáticos
    eRetos = 3                  ' Retos
    eRetosRapidos = 4       ' Retos Rapido
    eBuscadorAgites = 5     ' Buscador de Agites en ZONA INSEGURA
    eMixed = 6                  ' Realiza un MIXED con orden de prioridad.
    eUserList = 7               ' Busca según la lista de usuarios que decidió ser seguida por el BOT

    e_LAST = 8
End Enum

Public Type tStreamer
        Active As Integer ' Determina el Index del GM BOT (UserIndex)
        InitialPosition As WorldPos
        LastSummon As Long
        LastTarget As String
        UserIndex As Integer
        Last As Long
        Mode As eStreamerMode
        
        
        Config_TimeWarp As Long
        Config_TimeCanIndex As Long
        
        Users(1 To MAX_STREAM_USERS) As Integer     ' Usuarios que solicitaron al STREAMBOT
End Type

Public Const STREAMER_MAX_BOTS As Byte = 10
Public StreamerBot As tStreamer


' @ Inicializa al GM BOT, con una posición de respawn general.
Public Sub Streamer_Can(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Streamer_Can_Err
        '</EhHeader>
    
100     If Not EsGm(UserIndex) Then Exit Sub
    
102     With UserList(UserIndex)
104         If StreamerBot.Active Then
106             Streamer_Initial 0, 0, 0, 0
            Else
108             Streamer_Initial UserIndex, .Pos.Map, .Pos.X, .Pos.Y
            End If
        End With
    
        '<EhFooter>
        Exit Sub

Streamer_Can_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Can " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Streamer_Initial(ByVal UserIndex As Integer, _
                            ByVal Map As Integer, _
                            ByVal X As Byte, _
                            ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo Streamer_Initial_Err
        '</EhHeader>
    
        Dim tUser As Integer
            
100     With StreamerBot
              If Map > 0 Then
                If .Active > 0 Then
                  Call WriteConsoleMsg(UserIndex, "¡Está siendo utilizado por otro!", FontTypeNames.FONTTYPE_INFORED)
                  Exit Sub
                End If
              End If
              
102         .Active = UserIndex
104         .InitialPosition.Map = Map
106         .InitialPosition.X = X
108         .InitialPosition.Y = Y

            .Config_TimeWarp = STREAMER_TIME_AUTO_WARP
            .Config_TimeCanIndex = STREAMER_TIME_CAN_SEARCH
        
110         If .Active > 0 Then
112             Call Streamer_CheckPosition
            End If

        End With

        '<EhFooter>
        Exit Sub

Streamer_Initial_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Initial " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' @ Busca el UserIndex en EVENTOS AUTOMATICOS
Private Function Streamer_Search_Event(ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Streamer_Search_Event_Err
        '</EhHeader>

        Dim A          As Long, B As Long

        Dim BestTarget As Integer
        
100     For A = 1 To MAX_EVENT_SIMULTANEO

102         With Events(A)

104             If .Run Then
                    
106                 For B = LBound(.Users) To UBound(.Users)

108                     If .Users(B).ID > 0 Then
114                             If (Time - UserList(.Users(B).ID).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(.Users(B).ID).flags.Muerto = 0 Then
                                
116                                 Streamer_Search_Event = .Users(B).ID
118                                 StreamerBot.LastSummon = Time
120                                 UserList(.Users(B).ID).Counters.TimeGMBOT = Time
122                                 StreamerBot.LastTarget = UCase$(UserList(.Users(B).ID).Name)
                                    Exit Function
                                End If
                            End If

134                 Next B

                End If

            End With
    
136     Next A

        '<EhFooter>
        Exit Function

Streamer_Search_Event_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Event " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

 ' @ Busca el UserIndex en RETOS
Private Function Streamer_Search_Fight(ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Streamer_Search_Fight_Err
        '</EhHeader>

        Dim A As Long, B As Long
    

100     For A = 1 To MAX_RETOS_SIMULTANEOS

102         With Retos(A)

104             If .Run Then

106                 For B = LBound(.User) To UBound(.User)

108                     If .User(B).UserIndex > 0 Then
110                           If (Time - UserList(.User(B).UserIndex).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(.User(B).UserIndex).flags.Muerto = 0 Then
112                             Streamer_Search_Fight = .User(B).UserIndex
114                             StreamerBot.LastSummon = Time
116                             UserList(.User(B).UserIndex).Counters.TimeGMBOT = Time
118                             StreamerBot.LastTarget = UCase$(UserList(.User(B).UserIndex).Name)
                                Exit Function

                            End If
                    
                        End If

120                 Next B
            
                End If

            End With
    
122     Next A
        '<EhFooter>
        Exit Function

Streamer_Search_Fight_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Search_Fight " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function


 ' @ Busca el UserIndex en RETOS RAPIDOS
Private Function Streamer_Search_FightFast(ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Streamer_Search_FightFast_Err
        '</EhHeader>

        Dim A As Long, B As Long
    
100     For A = 1 To MAX_RETO_FAST

102         With RetoFast(A)
                If .Run Then
104                 For B = LBound(.Users) To UBound(.Users)
    
106                     If .Users(B) > 0 Then
108                             If (Time - UserList(.Users(B)).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(.Users(B)).flags.Muerto = 0 Then
110                             Streamer_Search_FightFast = .Users(B)
112                             StreamerBot.LastSummon = Time
114                             UserList(.Users(B)).Counters.TimeGMBOT = Time
116                             StreamerBot.LastTarget = UCase$(UserList(.Users(B)).Name)
                                Exit Function
    
                            End If
    
                        End If
    
118                 Next B
                End If
            End With
    
120     Next A
        '<EhFooter>
        Exit Function

Streamer_Search_FightFast_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Search_FightFast " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

' @ Buscamos usuarios en Zona Segura . @ NO WORKS
Private Function Streamer_Search_Secure(ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Streamer_Search_Secure_Err
        '</EhHeader>
        
        Dim A As Long
    
100     For A = 1 To LastUser

102         With UserList(A)

104             If (.ConnIDValida) Then
106                 If .flags.UserLogged Then
108                     If Not EsGm(A) Then
110                          If (Time - UserList(A).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex Then
112                             If Not MapInfo(.Pos.Map).Pk And UserList(A).flags.Muerto = 0 Then
114                                 Streamer_Search_Secure = A
116                                 StreamerBot.LastSummon = Time
                                      UserList(A).Counters.TimeGMBOT = Time
118                                 StreamerBot.LastTarget = UCase$(.Name)
                                    Exit Function

                                End If

                            End If

                        End If

                    End If

                End If
            
            End With

120     Next A
    
        '<EhFooter>
        Exit Function

Streamer_Search_Secure_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Search_Secure " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' @ Buscamos agites
Private Function Streamer_Search_Insecure(ByVal Time As Long) As Integer

        '<EhHeader>
        On Error GoTo Streamer_Search_Insecure_Err
        
        '</EhHeader>
        
        Dim A As Long
    
100     For A = 1 To LastUser

102         With UserList(A)

104             If (.ConnIDValida) Then
106                 If .flags.UserLogged Then
110                     If (Time - UserList(A).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex Then
112                         If MapInfo(.Pos.Map).Pk And MapInfo(.Pos.Map).NumUsers >= 7 And UserList(A).flags.Muerto = 0 Then
114                             Streamer_Search_Insecure = A
                                UserList(A).Counters.TimeGMBOT = Time
116                             StreamerBot.LastSummon = Time
118                             StreamerBot.LastTarget = UCase$(.Name)
                                Exit Function

                            End If

                        End If

                    End If

                End If
            
            End With

120     Next A
    
        '<EhFooter>
        Exit Function

Streamer_Search_Insecure_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Insecure_Err " & "at line " & Erl
        
        '</EhFooter>
End Function

' @ Buscamos uno de los usuarios que haya pedido ser seguido...
Private Function Streamer_Search_Users(ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Streamer_Search_Users_Err
        '</EhHeader>
        
        Dim A         As Long

        Dim UserIndex As Integer
    
100     For A = 1 To MAX_STREAM_USERS
102         UserIndex = StreamerBot.Users(A)

104         If UserIndex > 0 Then
            
106             If (Time - UserList(UserIndex).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(UserIndex).flags.Muerto = 0 Then
            
108                 Streamer_Search_Users = UserIndex
110                 UserList(UserIndex).Counters.TimeGMBOT = Time
112                 StreamerBot.LastSummon = Time
114                 StreamerBot.LastTarget = UCase$(UserList(UserIndex).Name)
                    Exit Function
                    
                End If

            End If
        
116     Next A
    
        '<EhFooter>
        Exit Function

Streamer_Search_Users_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Search_Users " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

' @ Agrega un usuario a la lista
Public Sub Streamer_RequiredBOT(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Streamer_RequiredBOT_Err
        '</EhHeader>
    
        Dim Slot As Byte
    
100     If StreamerBot.Active = 0 Then
102         Call WriteConsoleMsg(UserIndex, "El Hamster del CPU está descansando. Solicita nuestra pantalla LITOMANIA en otro momento.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
104     With UserList(UserIndex)
106         If .flags.BotList > 0 Then
108             Call WriteConsoleMsg(UserIndex, "¡Ya te encuentras en la lista de búsqueda del GM! Sal del Juego para no estarlo.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
        
110         Slot = Streamer_Required_Slot
        
112         If Slot > 0 Then
114             Call Streamer_SetBotList(UserIndex, Slot, False)
116             Call WriteConsoleMsg(UserIndex, "Te he agregado a mi lista... podrías ser el próximo ¡Asi que muestra algo o me iré!", FontTypeNames.FONTTYPE_INFOGREEN)
            Else
118             Call WriteConsoleMsg(UserIndex, "¡Vaya! Que solicitado soy... Espera un momento que renuevo la lista y vuelve a intentar pronto. Podré seguirte y ni te darás cuenta!", FontTypeNames.FONTTYPE_INFORED)
            
            End If
    
        End With
        '<EhFooter>
        Exit Sub

Streamer_RequiredBOT_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_RequiredBOT " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Streamer_SetBotList(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Killed As Boolean)
        '<EhHeader>
        On Error GoTo Streamer_SetBotList_Err
        '</EhHeader>

100     StreamerBot.Users(Slot) = IIf((Killed = True), 0, UserIndex)
102     UserList(UserIndex).flags.BotList = IIf((Killed = True), 0, Slot)
    
        '<EhFooter>
        Exit Sub

Streamer_SetBotList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_SetBotList " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' @ Busca un SLOT libre
Private Function Streamer_Required_Slot() As Byte
        '<EhHeader>
        On Error GoTo Streamer_Required_Slot_Err
        '</EhHeader>
        Dim A As Long
    
100     For A = 1 To MAX_STREAM_USERS
102         If StreamerBot.Users(A) = 0 Then
104             Streamer_Required_Slot = A
                Exit Function
            End If
106     Next A
    
        '<EhFooter>
        Exit Function

Streamer_Required_Slot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Required_Slot " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function


' @ Busca al proximo usuario disponible para tomar su posición
Public Function Streamer_Search_UserIndex(ByVal Time As Long) As Integer
        '<EhHeader>
        On Error GoTo Streamer_Search_UserIndex_Err
        '</EhHeader>

        Dim A As Long

        Dim B As Long
    
100     With StreamerBot
           
         
104         Select Case .Mode
        
                Case eStreamerMode.eZonaSegura
106                 Streamer_Search_UserIndex = Streamer_Search_Secure(Time)
                
108             Case eStreamerMode.eEventos
110                 Streamer_Search_UserIndex = Streamer_Search_Event(Time)
                
112             Case eStreamerMode.eRetos
114                 Streamer_Search_UserIndex = Streamer_Search_Fight(Time)
                
116             Case eStreamerMode.eRetosRapidos
118                 Streamer_Search_UserIndex = Streamer_Search_FightFast(Time)
                
120             Case eStreamerMode.eBuscadorAgites
                    ' @ Realizar comprobaciones de lanzamiento de hechizos y golpes
                      Streamer_Search_UserIndex = Streamer_Search_Insecure(Time)
                
                  Case eStreamerMode.eUserList
                     Streamer_Search_UserIndex = Streamer_Search_Users(Time)
                     
122             Case eStreamerMode.eMixed
                    ' Ordenar según la prioridad
124                 Streamer_Search_UserIndex = Streamer_Search_Event(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
126                 Streamer_Search_UserIndex = Streamer_Search_FightFast(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                    Streamer_Search_UserIndex = Streamer_Search_Fight(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                    Streamer_Search_UserIndex = Streamer_Search_Users(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                    Streamer_Search_UserIndex = Streamer_Search_Insecure(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                    Streamer_Search_UserIndex = Streamer_Search_Secure(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
            End Select
        End With
    
        '<EhFooter>
        Exit Function

Streamer_Search_UserIndex_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_Search_UserIndex " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Streamer_Sum(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    Dim X As Long
    Dim Y As Long
    
    With StreamerBot
    
        X = RandomNumber(UserList(UserIndex).Pos.X - 3, UserList(UserIndex).Pos.X + 3)
        Y = RandomNumber(UserList(UserIndex).Pos.Y - 3, UserList(UserIndex).Pos.Y + 3)
        Call EventWarpUser(.Active, UserList(UserIndex).Pos.Map, X, Y)
    End With
    
    Exit Sub
ErrHandler:
End Sub

' @ Reinicia cuando lo necesite
Public Sub Streamer_CheckUser(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim Time As Long
    
    With StreamerBot
    
        If .Active = 0 Then Exit Sub
        
        Time = GetTime
        
        If .UserIndex = UserIndex Then
            .UserIndex = Streamer_Search_UserIndex(Time)
            
            Call EventWarpUser(.Active, .InitialPosition.Map, .InitialPosition.X, .InitialPosition.Y)
            .LastSummon = Time
        End If
        
    End With
    
    Exit Sub
ErrHandler:
    
    
End Sub
' @ Comprueba que tan lejos se fue del objetivo
Public Sub Streamer_CheckPosition()
        '<EhHeader>
        On Error GoTo Streamer_CheckPosition_Err
        '</EhHeader>
               Dim Time      As Double
              Dim X As Integer, Y As Integer
              
         Time = GetTime
        
            
            Static SecondsCheckCercania As Integer
        
100     With StreamerBot
    
            ' @ El BOT no está activo.
102         If .Active = 0 Then Exit Sub
            
            
            SecondsCheckCercania = SecondsCheckCercania + 1
            
            
            If SecondsCheckCercania >= 2 Then
            
                If .UserIndex > 0 Then
                    With UserList(.UserIndex)
                         If Distance(UserList(StreamerBot.Active).Pos.X, UserList(StreamerBot.Active).Pos.Y, .Pos.X, .Pos.Y) >= 6 Then
                            Streamer_Sum StreamerBot.UserIndex
                        End If
                    End With
                End If
                
                SecondsCheckCercania = 0
            End If
             
            ' @ Segun el Tiempo SETEADO entre WARP & WARP
104         If (Time - .LastSummon) < .Config_TimeWarp Then Exit Sub

            ' @ Esto de abajo es llamado respetando cada 40s
            Dim UserIndex As Integer, Pos As WorldPos

108         UserIndex = Streamer_Search_UserIndex(Time)
            
110        If UserIndex > 0 Then
                  .UserIndex = UserIndex
112             Call EventWarpUser(.Active, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                .LastSummon = Time
            Else
                 .UserIndex = 0
                 
114             If Len(.LastTarget) = 0 Then
116                 If Distance(.InitialPosition.X, .InitialPosition.Y, UserList(StreamerBot.Active).Pos.X, UserList(StreamerBot.Active).Pos.Y) > 7 Then
118                     Call EventWarpUser(.Active, .InitialPosition.Map, .InitialPosition.X, .InitialPosition.Y)
                           .LastSummon = Time
                    End If

                End If
            
            End If
   
        End With
    
        '<EhFooter>
        Exit Sub

Streamer_CheckPosition_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mStreamer.Streamer_CheckPosition " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function Streamer_Mode_String(ByRef Mode As eStreamerMode) As String
        Select Case Mode
        
            Case eStreamerMode.eBuscadorAgites
                Streamer_Mode_String = "Agites en Zona Insegura"
            
            Case eStreamerMode.eEventos
                Streamer_Mode_String = "Eventos automáticos"
            
            Case eStreamerMode.eMixed
                Streamer_Mode_String = "Modalidad Mixed. Busca el mejor emparejamiento interno."
                
            Case eStreamerMode.eRetos
                Streamer_Mode_String = "Retos privados"
            
            Case eStreamerMode.eRetosRapidos
                Streamer_Mode_String = "Retos rapidos"
            
            Case eStreamerMode.eZonaSegura
                Streamer_Mode_String = "Usuarios en Zona segura. NO trabajadores."
        
        End Select
End Function
