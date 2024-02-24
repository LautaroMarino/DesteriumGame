Attribute VB_Name = "EventosDS"
' REFERENCIAS

'#################################
' EVENTO DE TELEPORTS
'#################################

'#################################
' EVENTO DE GRANBESTIA
'#################################

Option Explicit

' # Configuración General de los Eventos
Public Const MAX_EVENT_SIMULTANEO As Byte = 20
Public Const MAX_USERS_EVENT      As Byte = 64


' # Configuración de Arenas de Enfrentamiento
Public Const MAX_MAP_FIGHT        As Byte = 29
Public Const MAX_MAP_FIGHT_NORMAL        As Byte = 18
Public Const MAX_MAP_FIGHT_PLANTES       As Byte = 29



Public Const MAX_MAP_TELEPORT        As Byte = 6

' # Configuración de NPC Gran Bestia
Public Const NPC_GRAN_BESTIA      As Integer = 765

' # Tipos de Eventos disponibles
Public Enum eModalityEvent

    CastleMode = 1
    DagaRusa = 2
    DeathMatch = 3
    Enfrentamientos = 4
    Teleports = 5
    GranBestia = 6
    Busqueda = 7
    Unstoppable = 8
    JuegosDelHambre = 9
    
    Manual = 10
    
End Enum

Public Type tUserEvent

    ID As Integer
    Name As String
    Team As Byte
    Value As Integer
    Selected As Byte
    MapFight As Byte
    Oponent As Integer
    RoundsWin As Byte
    RoundsWinFinal As Integer
    
    TimeCancel As Long
    Damage As Long
End Type

Public Enum eFaction

    fCrim = 1
    fCiu = 2
    fLegion = 3
    fArmada = 4

End Enum

Public Enum eConfigEvent
    eBronce = 1
    ePlata = 2
    eOro = 3
    ePremium = 4
    eDañoZona = 5
    eAutoCupos = 6
    eInvFree = 7
    eParty = 8
    eGuild = 9
    eResu = 10
    eOcultar = 11
    eInvisibilidad = 12
    eInvocar = 13
    eMezclarApariencias = 14
    eDagaMaster = 15
    eSpellCuration = 16
    eUsePotion = 17
    eUseParalizar = 18
    eUseApocalipsis = 19
    eUseDescarga = 20
    eUseTormenta = 21
    eTeletransportacion = 22
    eCascoEscudo = 23
    eFuegoAmigo = 24
End Enum

Public Const MAX_EVENTS_CONFIG As Byte = 24
Public Const MAX_REWARD_OBJ As Byte = 10


Public Type EventsTime
    Fail As Byte
    
    InitDay As Byte ' Días de la semana (1-31)
    InitWeek As Byte ' Identificacion de días de la semana (Viernes,Lunes,Martes)
    InitHour As Byte ' Hora específica del Evento (18:00hs)
    InitMinute As Byte ' Hora específica del Evento (18:00hs)
End Type

Public Type tEvents
    Predeterminado As Boolean ' Determina si es un evento cargado desde un .ini, de manera tal que se tenga que reingresar al sistema y hacerse de forma ilimitada.
    TimeInit_Default As Long
    TimeCancel_Default As Long

    LastReward As Byte ' Ultimo objeto cargado
    RewardObj(1 To MAX_REWARD_OBJ) As Obj ' Lista de Premios Donados
    Name As String
    config(1 To MAX_EVENTS_CONFIG) As Byte
    Enabled As Boolean
    Run As Boolean
    Modality As eModalityEvent
    IsPlante As Byte
    TeamCant As Byte

    ArenasOcupadas As Byte
    ArenasLimit As Byte
    ArenasMin As Byte
    ArenasMax As Byte
    
    GanaSigue As Byte   ' Determina a cuantas victorias puede retirarse
    Quotas As Byte
    QuotasMin As Byte
    QuotasMax As Byte
    Inscribed As Byte
    Rounds As Byte
    
    LvlMax As Byte
    LvlMin As Byte
    
    InscriptionGld As Long
    InscriptionEldhir As Long
    
    AllowedClasses() As Byte
    AllowedFaction() As Byte
    
    InscriptionAcumulated As Byte
    PrizePoints As Integer
    PrizeEldhir As Integer
    PrizeGld As Long
    PrizeObj As Obj
    
    TempAdd As String
    TempDate As String
    TempFormat As String
    
    LimitRed As Integer
    LimitRound As Byte
    LimitRoundFinal As Byte
    
    TimeInscription As Long
    TimeCancel As Long
    TimeCount As Long
    TimeFinish As Long
    TimeInit As Long
    
    Users() As tUserEvent
    
    ' Por si alguno es con NPC
    NpcIndex As Integer
    
    ' Por si cambia el body del personaje y saca todo lo otro.
    CharHp As Integer
    
    npcUserIndex As Integer
    
    ChangeClass As Byte
    ChangeRaze As Byte
    ChangeLevel As Byte
    
    Time As EventsTime
    
    Prob As Byte            ' Si el evento o criatura del evento tiene una PROB se usa esto.
End Type


Public EventsCheck() As tEvents ' // Eventos predefinidos
Public EventLast As Integer
Public Events(1 To MAX_EVENT_SIMULTANEO) As tEvents

Private Enum eModeMap
        eDefault = 0    ' Arenas por Default
        ePlante = 1
End Enum

Private Type tMap

    Run As Boolean
    Map As Integer
    X As Byte
    Y As Byte
    MAP_TILE_VS As Byte
    
    ModeMap As eModeMap

End Type

Public Type tMapTeleport
    Usage As Boolean      ' Determina si se esta usando o no
    Map As Integer      ' Mapa que usa para el evento
    XWarp As Byte       ' Posición Warp X del User
    YWarp As Byte       ' Posición Warp Y del User
    
    XInitial_TP As Byte    ' Posición del primer tile apto Portal
    YInitial_TP As Byte    ' Posición del primer tile apto Portal
    
    XTiles_TP As Byte   ' Tiles en X que dura la Fila
    Y_Pasajes As Byte   ' Cantidad de Pasajes que tiene
    Y_TileAdd As Byte   ' Tiles Y que hay entre cada pasaje
    
    MaxQuotas As Byte   ' Máximo de usuarios permitidos en el evento
End Type

Public Type tMapEvent

    Fight(1 To MAX_MAP_FIGHT) As tMap
    Teleport() As tMapTeleport
    TeleportWin As tMap
    SalaEspera As tMap
    Castle(1) As tMap
    DagaRusa As tMap
    DeathMatch As tMap
    Busqueda As tMap
    Imparable As tMap
    JuegosDelHambre As tMap
End Type

Public MapEvent As tMapEvent
Public EventLastDefined As Integer


' @ Seguidilla de Eventossssssssssssssssss

Public SeguidillasLast As Byte

Public CopyEvents() As tEvents



' @
Public Sub LoadMapEvent()
    
    With MapEvent
        

            .Imparable.Map = 144
            .Imparable.X = 50
            .Imparable.Y = 50
        
            ' Coordenadas Random
            .Busqueda.Map = 136
            
            
            .JuegosDelHambre.Map = 140
            '.JuegosDelHambre.X = 9
            '.JuegosDelHambre.y = 86
            
            '.JuegosDelHambre.Map = 175
            '.JuegosDelHambre.Map = 176
             
            ' Chica
            '.Teleport(1).Map = 142
            '.Teleport(1).X = 25
            '.Teleport(1).Y = 57
        
            ' Chica
           ' .Teleport.Map = 142
            '.Teleport.X = 77
            '.Teleport.Y = 57
            
            ' Teleports medio
            ' .Teleport.Map = 163
            '.Teleport.X = 22
            '.Teleport.Y = 61
            '.Teleport.X = 76
            '.Teleport.Y = 61
            
            ' Teleports grande
            ' .Teleport.Map = 164
            '.Teleport.X = 27
            '.Teleport.Y = 81
            '.Teleport.X = 72
            '.Teleport.Y = 81
            
            .DeathMatch.Map = 135
            .DeathMatch.X = 70
            .DeathMatch.Y = 30
        
            .DagaRusa.Map = 135
            .DagaRusa.X = 21
            .DagaRusa.Y = 60
        
            .Castle(0).Map = 141
            .Castle(0).X = 50
            .Castle(0).Y = 21
        
            .Castle(1).Map = 141
            .Castle(1).X = 50
            .Castle(1).Y = 71
        
            .SalaEspera.Map = 58
            .SalaEspera.X = 28
            .SalaEspera.Y = 20
        
            .Fight(1).Run = False
            .Fight(1).Map = 73
            .Fight(1).X = 16
            .Fight(1).Y = 12
            .Fight(1).MAP_TILE_VS = 16
            
            .Fight(2).Run = False
            .Fight(2).Map = 73
            .Fight(2).X = 16
            .Fight(2).Y = 41
            .Fight(2).MAP_TILE_VS = 16
            
            .Fight(3).Run = False
            .Fight(3).Map = 73
            .Fight(3).X = 16
            .Fight(3).Y = 68
            .Fight(3).MAP_TILE_VS = 16
            
            .Fight(4).Run = False
            .Fight(4).Map = 73
            .Fight(4).X = 46
            .Fight(4).Y = 12
            .Fight(4).MAP_TILE_VS = 16
            
            .Fight(5).Run = False
            .Fight(5).Map = 73
            .Fight(5).X = 46
            .Fight(5).Y = 41
            .Fight(5).MAP_TILE_VS = 16
                        
            .Fight(6).Run = False
            .Fight(6).Map = 73
            .Fight(6).X = 46
            .Fight(6).Y = 68
            .Fight(6).MAP_TILE_VS = 16
            
            .Fight(7).Run = False
            .Fight(7).Map = 74
            .Fight(7).X = 16
            .Fight(7).Y = 12
            .Fight(7).MAP_TILE_VS = 16
               
            .Fight(8).Run = False
            .Fight(8).Map = 74
            .Fight(8).X = 16
            .Fight(8).Y = 41
            .Fight(8).MAP_TILE_VS = 16
             
            .Fight(9).Run = False
            .Fight(9).Map = 74
            .Fight(9).X = 16
            .Fight(9).Y = 68
            .Fight(9).MAP_TILE_VS = 16
              
            .Fight(10).Run = False
            .Fight(10).Map = 74
            .Fight(10).X = 46
            .Fight(10).Y = 12
            .Fight(10).MAP_TILE_VS = 16
            
            .Fight(11).Run = False
            .Fight(11).Map = 74
            .Fight(11).X = 46
            .Fight(11).Y = 41
            .Fight(11).MAP_TILE_VS = 16
            
            .Fight(12).Run = False
            .Fight(12).Map = 74
            .Fight(12).X = 46
            .Fight(12).Y = 68
            .Fight(12).MAP_TILE_VS = 16
            
            .Fight(13).Run = False
            .Fight(13).Map = 75
            .Fight(13).X = 16
            .Fight(13).Y = 12
            .Fight(13).MAP_TILE_VS = 16
            
            .Fight(14).Run = False
            .Fight(14).Map = 75
            .Fight(14).X = 16
            .Fight(14).Y = 41
            .Fight(14).MAP_TILE_VS = 16
            
            .Fight(15).Run = False
            .Fight(15).Map = 75
            .Fight(15).X = 16
            .Fight(15).Y = 68
            .Fight(15).MAP_TILE_VS = 16
            
            .Fight(16).Run = False
            .Fight(16).Map = 75
            .Fight(16).X = 46
            .Fight(16).Y = 12
            .Fight(16).MAP_TILE_VS = 16
            
            .Fight(17).Run = False
            .Fight(17).Map = 75
            .Fight(17).X = 46
            .Fight(17).Y = 41
            .Fight(17).MAP_TILE_VS = 16
            
            .Fight(18).Run = False
            .Fight(18).Map = 75
            .Fight(18).X = 46
            .Fight(18).Y = 68
            .Fight(18).MAP_TILE_VS = 16
            
            ' Arenas de Plante
            .Fight(19).Run = False
            .Fight(19).Map = 76
            .Fight(19).X = 18
            .Fight(19).Y = 16
            .Fight(19).ModeMap = ePlante
            .Fight(19).MAP_TILE_VS = 1
             
            .Fight(20).Run = False
            .Fight(20).Map = 76
            .Fight(20).X = 18
            .Fight(20).Y = 31
            .Fight(20).ModeMap = ePlante
            .Fight(20).MAP_TILE_VS = 1
            
            .Fight(21).Run = False
            .Fight(21).Map = 76
            .Fight(21).X = 18
            .Fight(21).Y = 45
            .Fight(21).ModeMap = ePlante
            .Fight(21).MAP_TILE_VS = 1
            
            .Fight(22).Run = False
            .Fight(22).Map = 76
            .Fight(22).X = 18
            .Fight(22).Y = 59
            .Fight(22).ModeMap = ePlante
            .Fight(22).MAP_TILE_VS = 1

            .Fight(23).Run = False
            .Fight(23).Map = 76
            .Fight(23).X = 18
            .Fight(23).Y = 72
            .Fight(23).ModeMap = ePlante
            .Fight(23).MAP_TILE_VS = 1
            
            .Fight(24).Run = False
            .Fight(24).Map = 76
            .Fight(24).X = 18
            .Fight(24).Y = 85
            .Fight(24).ModeMap = ePlante
            .Fight(24).MAP_TILE_VS = 1
            
            .Fight(25).Run = False
            .Fight(25).Map = 76
            .Fight(25).X = 51
            .Fight(25).Y = 16
            .Fight(25).ModeMap = ePlante
            .Fight(25).MAP_TILE_VS = 1
            
            .Fight(26).Run = False
            .Fight(26).Map = 76
            .Fight(26).X = 51
            .Fight(26).Y = 31
            .Fight(26).ModeMap = ePlante
            .Fight(26).MAP_TILE_VS = 1
            
            .Fight(27).Run = False
            .Fight(27).Map = 76
            .Fight(27).X = 51
            .Fight(27).Y = 45
            .Fight(27).ModeMap = ePlante
            .Fight(27).MAP_TILE_VS = 1
            
            .Fight(28).Run = False
            .Fight(28).Map = 76
            .Fight(28).X = 51
            .Fight(28).Y = 59
            .Fight(28).ModeMap = ePlante
            .Fight(28).MAP_TILE_VS = 1
            
            .Fight(29).Run = False
            .Fight(29).Map = 76
            .Fight(29).X = 51
            .Fight(29).Y = 72
            .Fight(29).ModeMap = ePlante
            .Fight(29).MAP_TILE_VS = 1

            .TeleportWin.Map = 142
            .TeleportWin.X = 25
            .TeleportWin.Y = 15
       
    End With

    ' Leemos los eventos por defecto
    Call Events_Read

End Sub


' Eventos 100% Autonomos
Public Sub Events_Loop_Check()
        '<EhHeader>
        On Error GoTo Events_Loop_Check_Err
        '</EhHeader>
    
        Dim A As Long
    
100     For A = 1 To EventLastDefined

102         With EventsCheck(A)

104            If .Run = False Then
106                 Call Events_CheckTime(A)
                End If

            End With
        
108     Next A

        '<EhFooter>
        Exit Sub

Events_Loop_Check_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_Loop_Check " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Events_CheckTime(ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo Events_CheckTime_Err
        '</EhHeader>

        Dim Time As Date
    
100     Time = Now

102     With EventsCheck(SlotEvent).Time
        
            'EXAMPLE:  1° de cada mes
104         If .InitDay > 0 Then
106             If Day(Time) <> .InitDay Then Exit Sub
            End If
        
            ' Días: SABADO
108         If .InitWeek > 0 Then
110             If Weekday(Time) <> .InitWeek Then Exit Sub
            End If
        
            ' Horario específico: 18hs
112         If .InitHour > 0 Then
114             If Hour(Time) <> .InitHour Then Exit Sub
            End If
        
            ' Minuto específico: 18:30hs . Capaz se desfasó porque otro de igual categoría estaba en curso, entonces se sigue intentando poner pronto (proximos minutos tolerancia)
116         If .InitMinute > 0 And .Fail = 0 Then
118             If Minute(Time) <> .InitMinute Then Exit Sub
            End If
        
            Dim Slot As Byte
120         Slot = NewEvent(EventsCheck(SlotEvent))
        
122         If Slot = 0 Then
124             .Fail = .Fail + 1
            
126             If .Fail = 10 Then .Fail = 0
            Else
128             EventsCheck(SlotEvent).Run = True
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

Events_CheckTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_CheckTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
    
Public Sub Events_Read()
        '<EhHeader>
        On Error GoTo Events_Read_Err
        '</EhHeader>
        Dim Read      As clsIniManager, A As Long, B As Long, Temp As String
    
        Dim TempEvent As tEvents
100     Set Read = New clsIniManager
    
102     Call Read.Initialize(App.Path & "\dat\events\events_list.ini")
    
104     EventLastDefined = val(Read.GetValue("INIT", "LAST"))
    
106     ReDim EventsCheck(0 To EventLastDefined) As tEvents
    
108     For A = 1 To EventLastDefined
110         With EventsCheck(A)
        
112             ReDim .AllowedClasses(1 To NUMCLASES) As Byte
114             ReDim .AllowedFaction(1 To 4) As Byte

116             .Name = Read.GetValue(A, "NAME")
118             .Quotas = Read.GetValue(A, "QUOTAS")
120             .TeamCant = Read.GetValue(A, "TEAMCANT")
122             .Modality = val(Read.GetValue(A, "MODALITY"))

126             .PrizeGld = val(Read.GetValue(A, "REWARDGLD"))
128             .LimitRound = val(Read.GetValue(A, "ROUNDS"))
130             .LimitRoundFinal = val(Read.GetValue(A, "ROUNDSFINAL"))
            
132             .config(eConfigEvent.eParty) = val(Read.GetValue(A, "CONPARTY"))
134             .config(eConfigEvent.eBronce) = val(Read.GetValue(A, "ITEMBRONCE"))
136             .config(eConfigEvent.ePlata) = val(Read.GetValue(A, "ITEMPLATA"))
138             .config(eConfigEvent.eOro) = val(Read.GetValue(A, "ITEMORO"))
140             .config(eConfigEvent.ePremium) = val(Read.GetValue(A, "ITEMPREMIUM"))
            
142             .config(eConfigEvent.eUseParalizar) = val(Read.GetValue(A, "UseParalizar"))
144             .config(eConfigEvent.eUseApocalipsis) = val(Read.GetValue(A, "UseApocalipsis"))
146             .config(eConfigEvent.eUseDescarga) = val(Read.GetValue(A, "UseDescarga"))
148             .config(eConfigEvent.eUseTormenta) = val(Read.GetValue(A, "UseTormenta"))
            
150             Temp = Read.GetValue(A, "LVL")
152             .LvlMin = val(ReadField(1, Temp, 45))
154             .LvlMax = val(ReadField(2, Temp, 45))
            
156             Temp = Read.GetValue(A, "REWARDOBJ")
            
158             .PrizeObj.ObjIndex = val(ReadField(1, Temp, 45))
160             .PrizeObj.Amount = val(ReadField(2, Temp, 45))
            
                ' Día Especifico (número)
                ' Día Específico (lunes, martes, miercoles, sabado)
                ' Horario Específico (18:00hs)
162             .Time.InitDay = val(Read.GetValue(A, "INITDAY"))
164             .Time.InitWeek = val(Read.GetValue(A, "INITWEEK"))
166             .Time.InitHour = val(Read.GetValue(A, "INITHOUR"))
168             .Time.InitMinute = val(Read.GetValue(A, "INITMINUTE"))
            
            
170             Temp = Read.GetValue(A, "Class")
                Dim TempClass As eClass
            
                Dim ArrayTemp() As String
172             ArrayTemp = Split(Read.GetValue(A, "Class"), "-")
            
174             For B = LBound(ArrayTemp) To UBound(ArrayTemp)
176                 .AllowedClasses(val(ArrayTemp(B))) = 1
178             Next B
            
180             For B = 1 To 4
182                 .AllowedFaction(B) = 1
184             Next B
            
186             For B = 1 To MAX_REWARD_OBJ
188                 Temp = Read.GetValue(A & "-REWARD", B)
                
190                 .RewardObj(B).ObjIndex = val(ReadField(1, Temp, 45))
192                 .RewardObj(B).Amount = val(ReadField(2, Temp, 45))
                
194                 If .RewardObj(B).ObjIndex > 0 Then .LastReward = .LastReward + 1

196             Next B
            
            End With
    
198     Next A
    
200     Set Read = Nothing
        '<EhFooter>
        Exit Sub

Events_Read_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_Read " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' Guarda el tiempo para que el evento se haga de forma autom?tica
Public Sub Events_Data_Predetermined()
        '<EhHeader>
        On Error GoTo Events_Data_Predetermined_Err
        '</EhHeader>
        Dim Manager As clsIniManager, A As Long, B As Long
    
100     Set Manager = New clsIniManager
    
102     Call Manager.Initialize(App.Path & "\dat\events\events_list.ini")
    
104     For A = 1 To EventLastDefined
106         With Events(A)
            
108             For B = 1 To MAX_REWARD_OBJ
110                  Call Manager.ChangeValue(A & "-REWARD", B, CStr(.RewardObj(B).ObjIndex & "-" & CStr(.RewardObj(B).Amount)))
112             Next B
            
            End With
114     Next A
    
116     Call Manager.DumpFile(App.Path & "\dat\events\events_list.ini")
    
118     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

Events_Data_Predetermined_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_Data_Predetermined " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Chequea si hay un mapa disponible antes de lanzar el acortamiento de cupos.. (Evento TELEPORTS)
Private Sub Events_CheckInitial_Map(ByVal SlotEvent As Byte)
    With Events(SlotEvent)
        
        Dim MapIndex As Integer
        
        MapIndex = Events_Teleports_SearchMapFree(.Inscribed)
        
        If MapIndex > 0 Then
            Call Events_UpdateQuotas(SlotEvent)
        End If
        
    End With
End Sub
'/MANEJO DE LOS TIEMPOS '/
Public Sub LoopEvent()

    On Error GoTo error

    Dim LoopC As Long

    Dim LoopY As Integer
    
    Dim Time As Long
    
    Dim A As Long
    
    For LoopC = 1 To MAX_EVENT_SIMULTANEO

        With Events(LoopC)

            If .Enabled Then
               
               For A = LBound(.Users) To UBound(.Users)
                    If .Users(A).ID > 0 Then
                        
                        If .Users(A).TimeCancel > 0 Then
                        
                            .Users(A).TimeCancel = .Users(A).TimeCancel - 1
                            
                            If .Users(A).TimeCancel <= 0 Then
                                Call WriteConsoleMsg(.Users(A).ID, "Tu duelo deberia estar siendo cancelado...", FontTypeNames.FONTTYPE_INFORED)
                                Call SendData(SendTarget.ToOne, .Users(A).ID, PrepareMessageRenderConsole("Daño total causado: " & .Users(A).Damage, eDamageType.d_DañoUserSpell, 10000, 0))
                                
                                ' # Check tiempo jugado
                            
                            End If

                        End If
                    End If
               
               Next A
               
                ' Cada 180 segundos (3 minutos) se comprueba si se puede completar el evento para que sea enviado.
                If (.TimeCancel > 0) Then
                    .TimeCancel = .TimeCancel - 1
                    
                    If .TimeCancel = 0 Then
                        Call CloseEvent(LoopC, , True)
                    Else
                        If (.TimeCancel Mod 60 = 0) Then
                            Select Case .Modality
                            
                                Case eModalityEvent.Teleports
                                    Call Events_CheckInitial_Map(LoopC)
                                    
                                Case Else
                                
                                    Call Events_UpdateQuotas(LoopC)
                            End Select
                            
                            Call SendData(SendTarget.toMap, 1, PrepareMessageRenderConsole(MODALITY_STRING(.Modality, .TeamCant, .IsPlante) & "» En 30 segundos se volvera a controlar cupos para poder comenzar.", d_AddGld, 6000, 0))
                          
                        End If
                    End If
                End If
                      
                ' Tiempo restante para que un evento comience
                If .TimeInit > 0 Then
                    .TimeInit = .TimeInit - 1
                            
                    If .TimeInit <= 0 Then
                        Call InitEvent(LoopC)
                    End If
                End If
                      
                ' Cuenta regresiva de eventos [SACAR]
                If .TimeCount > 0 Then
                    .TimeCount = .TimeCount - 1
                          
                    For LoopY = LBound(.Users()) To UBound(.Users())

                        If .Users(LoopY).ID > 0 Then
                            If .TimeCount = 0 Then
                                WriteConsoleMsg .Users(LoopY).ID, "Cuenta» ¡Comienza!", FontTypeNames.FONTTYPE_FIGHT
                                'WriteShortMsj .Users(LoopY).Id, 31, FontTypeNames.FONTTYPE_FIGHT
                            Else
                                WriteConsoleMsg .Users(LoopY).ID, "Cuenta» " & .TimeCount, FontTypeNames.FONTTYPE_EVENT
                                'WriteShortMsj .Users(LoopY).Id, 32, FontTypeNames.FONTTYPE_GUILD, .TimeCount
                            End If
                        End If

                    Next LoopY

                End If
                      
                ' Tiempo restante para que un evento finalice de forma automática.
                If .TimeFinish > 0 Then
                    .TimeFinish = .TimeFinish - 1
                          
                    If .TimeFinish = 0 Then
                        Call Events_FinishTime(LoopC)
                    End If
                End If
                
            End If
          
        End With

    Next LoopC
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : LoopEvent()"
End Sub

'// Funciones generales '//
Private Function FreeSlotEvent() As Byte
        '<EhHeader>
        On Error GoTo FreeSlotEvent_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     For LoopC = 1 To MAX_EVENT_SIMULTANEO

102         If Not Events(LoopC).Enabled Then
104             FreeSlotEvent = LoopC

                Exit For

            End If

106     Next LoopC

        '<EhFooter>
        Exit Function

FreeSlotEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.FreeSlotEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function Event_ModalityRepeat(ByVal Modality As eModalityEvent) As Boolean
        '<EhHeader>
        On Error GoTo Event_ModalityRepeat_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     For LoopC = 1 To MAX_EVENT_SIMULTANEO

102         If Events(LoopC).Modality = Modality Then
104             Event_ModalityRepeat = True

                Exit Function

            End If

106     Next LoopC

        '<EhFooter>
        Exit Function

Event_ModalityRepeat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_ModalityRepeat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Private Function Event_NameRepeat(ByVal Name As String) As Boolean
        '<EhHeader>
        On Error GoTo Event_NameRepeat_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     For LoopC = 1 To MAX_EVENT_SIMULTANEO

102         If StrComp(UCase$(Events(LoopC).Name), UCase$(Name)) = 0 Then
104             Event_NameRepeat = True

                Exit Function

            End If

106     Next LoopC

        '<EhFooter>
        Exit Function

Event_NameRepeat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_NameRepeat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Private Function GenerateSuffix(ByVal Index As Integer) As String
    ' Los sufijos que queremos usar
    Dim suffixes As String
    suffixes = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    ' Si el índice es mayor que la cantidad de sufijos, devuelve una cadena vacía
    If Index > Len(suffixes) Then
        GenerateSuffix = ""
    Else
        ' De lo contrario, devuelve el sufijo correspondiente
        GenerateSuffix = mid(suffixes, Index, 1)
    End If
End Function
Private Function FreeSlotUser(ByVal SlotEvent As Byte) As Byte
        '<EhHeader>
        On Error GoTo FreeSlotUser_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     With Events(SlotEvent)

102         For LoopC = 1 To MAX_USERS_EVENT

104             If .Users(LoopC).Name = vbNullString Then
106                 FreeSlotUser = LoopC

                    Exit Function

                End If

108         Next LoopC

        End With
          
        '<EhFooter>
        Exit Function

FreeSlotUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.FreeSlotUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function FreeSlotArena(ByVal SlotEvent As Byte) As Byte

        '<EhHeader>
        On Error GoTo FreeSlotArena_Err

        '</EhHeader>

        Dim LoopC            As Integer

        Dim Slot             As Byte

        Dim Searched         As Boolean

        Dim Temp             As Byte

        Dim IntentosFallidos As Byte
        
100     FreeSlotArena = 0
          
102     With Events(SlotEvent)

104         Do While Not Searched
106             Temp = RandomNumber(.ArenasMin, .ArenasMax)
 
108             If MapEvent.Fight(Temp).Run = False And .ArenasOcupadas < .ArenasLimit Then
110                 .ArenasOcupadas = .ArenasOcupadas + 1
112                 Searched = True
                    Exit Do
                Else
114                 IntentosFallidos = IntentosFallidos + 1
                
116                 If IntentosFallidos >= 4 Then
                        Exit Function

                    End If

                End If
             
            Loop
         
118         FreeSlotArena = Temp
          
        End With

        '<EhFooter>
        Exit Function

FreeSlotArena_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.EventosDS.FreeSlotArena " & "at line " & Erl

        

        '</EhFooter>
End Function

Public Function strUsersEvent(ByVal SlotEvent As Byte) As String
        '<EhHeader>
        On Error GoTo strUsersEvent_Err
        '</EhHeader>

        ' Texto que marca los personajes que están en el evento.
        Dim LoopC As Integer
          
100     With Events(SlotEvent)

102         For LoopC = LBound(.Users()) To UBound(.Users())

104             If .Users(LoopC).ID > 0 Then
106                 strUsersEvent = strUsersEvent & UserList(.Users(LoopC).ID).Name & "-"
                Else
108                 strUsersEvent = strUsersEvent & "(Vacio)" & "-"
                End If

110         Next LoopC

        End With

        '<EhFooter>
        Exit Function

strUsersEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.strUsersEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function CheckAllowedClasses(ByRef AllowedClasses() As Byte) As String
        '<EhHeader>
        On Error GoTo CheckAllowedClasses_Err
        '</EhHeader>

        Dim LoopC As Integer

100     Dim Valid As Boolean: Valid = True
    
102     For LoopC = 1 To NUMCLASES

104         If AllowedClasses(LoopC) = 1 Then
106             If CheckAllowedClasses = vbNullString Then
108                 CheckAllowedClasses = ListaClases(LoopC)
                Else
110                 CheckAllowedClasses = CheckAllowedClasses & ", " & ListaClases(LoopC)
                End If

            Else
112             Valid = False
            End If

114     Next LoopC

116     If Valid Then
118         CheckAllowedClasses = "TODAS"
        End If

        '<EhFooter>
        Exit Function

CheckAllowedClasses_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.CheckAllowedClasses " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function Events_CheckAllowed_Faction(ByRef AllowedFaction() As Byte) As String
        '<EhHeader>
        On Error GoTo Events_CheckAllowed_Faction_Err
        '</EhHeader>

        Dim LoopC As Integer

100     Dim Valid As Boolean: Valid = True
    
102     For LoopC = 1 To 4

104         If AllowedFaction(LoopC) = 1 Then
106             Events_CheckAllowed_Faction = Events_CheckAllowed_Faction & Faction_String(LoopC) & ", "
            Else
108             Valid = False
            End If

110     Next LoopC

112     If Len(Events_CheckAllowed_Faction) > 0 Then
114         Events_CheckAllowed_Faction = Left$(Events_CheckAllowed_Faction, Len(Events_CheckAllowed_Faction) - 2)
        End If

116     If Valid Then
118         Events_CheckAllowed_Faction = "TODAS"
        End If

        '<EhFooter>
        Exit Function

Events_CheckAllowed_Faction_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_CheckAllowed_Faction " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Events_SearchSlotEvent(ByVal Modality As String) As Byte
        '<EhHeader>
        On Error GoTo Events_SearchSlotEvent_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     Events_SearchSlotEvent = 0
          
102     For LoopC = 1 To MAX_EVENT_SIMULTANEO

104         With Events(LoopC)
106             If StrComp(UCase$(.Name), UCase$(Modality)) = 0 Then
108                 Events_SearchSlotEvent = LoopC

                    Exit Function

                End If
            
            End With

110     Next LoopC

        '<EhFooter>
        Exit Function

Events_SearchSlotEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_SearchSlotEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub EventWarpUser(ByVal UserIndex As Integer, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)

    ' // NUEVO
    
    On Error GoTo error

    ' Teletransportamos a cualquier usuario que cumpla con la regla de estar en un evento.
          
    Dim Pos As WorldPos
          
    With UserList(UserIndex)
        Pos.Map = Map
        Pos.X = X
        Pos.Y = Y
              
        ClosestStablePos Pos, Pos
        WarpUserChar UserIndex, Pos.Map, Pos.X, Pos.Y, False
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : EventWarpUser()"
End Sub

Private Sub ResetEvent(ByVal Slot As Byte)

    On Error GoTo error

    Dim LoopC As Integer
    Dim Predeterminado As Boolean
    Dim NullEvent As tEvents
    
    With Events(Slot)
         
        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                AbandonateEvent .Users(LoopC).ID, False, , True
            End If

        Next LoopC
              
        If .NpcIndex > 0 Then Call QuitarNPC(.NpcIndex)
    
    End With
    
    EventLast = EventLast - 1
    Events(Slot) = NullEvent
    
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : ResetEvent()"
End Sub

Public Function Events_CheckUserEvent(ByVal UserIndex As Integer, _
                                      ByVal SlotEvent As Byte, _
                                      ByRef ErrorMsg As String) As Boolean

    On Error GoTo error

    Events_CheckUserEvent = False
              
    With UserList(UserIndex)

        If .flags.Muerto Then
            ErrorMsg = "Estás muerto."

            Exit Function

        End If

        If .flags.Mimetizado Then
            ErrorMsg = "Estás mimetizado."

            Exit Function

        End If
              
        If .flags.Navegando Then
            ErrorMsg = "Estás Navegando."

            Exit Function

        End If
              
        If .flags.Desafiando > 0 Then
            ErrorMsg = "Estás desafiando."
                
            Exit Function

        End If
            
        If .flags.Invisible Then
            ErrorMsg = "Estás invisible."

            Exit Function

        End If
              
        If .flags.SlotEvent > 0 Then
            ErrorMsg = "Estás en Evento."

            Exit Function

        End If
              
        If .flags.SlotReto > 0 Or .flags.SlotFast > 0 Then
            ErrorMsg = "Estás en Reto."

            Exit Function

        End If
              
        If .Counters.Pena > 0 Then
            ErrorMsg = "Estás en la carcel."

            Exit Function

        End If
            
        If Not Is_Map_valid(UserIndex) Then
            ErrorMsg = "Estás en un mapa inválido."

            Exit Function

        End If
              
        If .flags.Comerciando Then
            ErrorMsg = "Estás comerciando."

            Exit Function

        End If
        
        If .flags.Desnudo = 1 And Events(SlotEvent).config(eConfigEvent.eInvFree) = 0 Then
            ErrorMsg = "¡No puedes entrar desnudo!"

            Exit Function
        
        End If
              
        If .Stats.MinAGU <= (.Stats.MaxAGU / 2) Or .Stats.MinHam <= (.Stats.MaxHam / 2) Then
            ErrorMsg = "¡Deberias comer y beber algo más para asegurarte de poder pasar todo el evento!"

            Exit Function
        
        End If
              
        If Events(SlotEvent).Run Then
            ErrorMsg = "El evento ya completo los cupos. Mejor suerte para la próxima."

            Exit Function

        End If
        
        If Events(SlotEvent).Quotas > 0 Then
            If Events(SlotEvent).Inscribed = Events(SlotEvent).Quotas Then
                ErrorMsg = "El evento ya completo los cupos. Mejor suerte para la próxima."
    
                Exit Function
            
            End If

        End If
        
        If Events(SlotEvent).LvlMin <> 0 Then
            If Events(SlotEvent).LvlMin > .Stats.Elv Then
                ErrorMsg = "Tu nivel no te permite entrar al evento."

                Exit Function

            End If

        End If
              
        If Events(SlotEvent).LvlMin <> 0 Then
            If Events(SlotEvent).LvlMax < .Stats.Elv Then
                ErrorMsg = "Tu nivel no te permite entrar al evento."

                Exit Function

            End If

        End If
              
        If Events(SlotEvent).AllowedClasses(.Clase) = 0 Then
            ErrorMsg = "Tu clase no te permite entrar al evento."

            Exit Function

        End If
              
        If Events(SlotEvent).InscriptionGld > .Stats.Gld Then
            ErrorMsg = "No tienes suficientes Monedas de Oro."

            Exit Function

        End If
              
        If Events(SlotEvent).InscriptionEldhir > .Stats.Eldhir Then
            ErrorMsg = "No tienes suficientes Monedas Dsp."

            Exit Function

        End If

        'If Events(SlotEvent).Config(eConfigEvent.eInvFree) = 1 Then
            'If .Invent.NroItems > 0 Then
                'ErrorMsg = "Debes tener el inventario vacío para poder participar de este evento"

                'Exit Function

            'End If

       'End If

        ' NO permitimos objetos especiales?¿
        Dim ErrorItem As String

        ErrorItem = TieneObjetos_Especiales(UserIndex, Events(SlotEvent).config(eConfigEvent.eBronce), Events(SlotEvent).config(eConfigEvent.ePlata), Events(SlotEvent).config(eConfigEvent.eOro), Events(SlotEvent).config(eConfigEvent.ePremium))
        
        If ErrorItem <> vbNullString Then
            ErrorMsg = ErrorItem
                    
            Exit Function

        End If
            
    End With

    Events_CheckUserEvent = True
          
    Exit Function

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Events_CheckUserEvent()"

End Function

Public Function MODALITY_STRING(ByVal Modality As eModalityEvent, _
                                ByVal TeamCant As Byte, _
                                ByVal IsPlante As Byte) As String
    Dim suffix As String
    Dim suffixIndex As Long
    suffixIndex = 1

    Select Case Modality
        Case eModalityEvent.CastleMode
            MODALITY_STRING = "REYVSREY"
        Case eModalityEvent.DagaRusa
            MODALITY_STRING = "DAGARUSA"
        Case eModalityEvent.DeathMatch
            MODALITY_STRING = "DEATHMATCH"
        Case eModalityEvent.Enfrentamientos
            MODALITY_STRING = IIf(IsPlante = 1, "PLANTE", TeamCant & "vs" & TeamCant)
        Case eModalityEvent.Enfrentamientos
            MODALITY_STRING = "MANUAL"
        Case eModalityEvent.Teleports
            MODALITY_STRING = "TELEPORTS"
        Case eModalityEvent.GranBestia
            MODALITY_STRING = "GRANBESTIA"
        Case eModalityEvent.JuegosDelHambre
            MODALITY_STRING = "JUEGOSDELHAMBRE"
        Case eModalityEvent.Busqueda
            MODALITY_STRING = "BUSQUEDA"
        Case eModalityEvent.Unstoppable
            MODALITY_STRING = "IMPARABLE"
    End Select

    ' Comprueba si el nombre del evento ya existe, si es así, entra en el bucle para añadir sufijos
    If Event_NameRepeat(MODALITY_STRING) Then
        ' Genera sufijos y comprueba si el nombre con el sufijo ya existe
        Do
            suffix = GenerateSuffix(suffixIndex)
            suffixIndex = suffixIndex + 1
        Loop While Event_NameRepeat(MODALITY_STRING & suffix)
    End If

    ' Añade el sufijo al nombre del evento
    MODALITY_STRING = MODALITY_STRING & suffix

End Function

' Modalidad de un reto
Public Function strModality(ByVal SlotEvent As Byte, _
                            ByVal Modality As eModalityEvent) As String
        '<EhHeader>
        On Error GoTo strModality_Err
        '</EhHeader>

100     strModality = Events(SlotEvent).Name

        '<EhFooter>
        Exit Function

strModality_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.strModality " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function strDescEvent(ByVal SlotEvent As Byte, _
                              ByVal Modality As eModalityEvent) As String
        '<EhHeader>
        On Error GoTo strDescEvent_Err
        '</EhHeader>

        ' Descripción del evento en curso.
100     Select Case Modality

            Case eModalityEvent.CastleMode
102             strDescEvent = "» Los usuarios entrarán de forma aleatorea para formar dos equipos. Ambos equipos deberán defender a su rey y a su vez atacar al del equipo contrario."

104         Case eModalityEvent.DagaRusa
106             strDescEvent = "» Los usuarios se teletransportarán a una posición donde estará un asesino dispuesto a apuñalarlos y acabar con su vida. El último que quede en pie es el ganador del evento."

108         Case eModalityEvent.DeathMatch
110             strDescEvent = "» Los usuarios ingresan y luchan en una arena donde se toparan con todos los demás concursantes. El que logre quedar en pie, será el ganador."

112         Case eModalityEvent.Busqueda
114             strDescEvent = "» Los personajes son teletransportados en un mapa donde su función principal será la recolección de objetos en el piso, para que así luego de tres minutos, el que recolecte más, ganará el evento."

116         Case eModalityEvent.Unstoppable
118             strDescEvent = "» Los personajes lucharan en un TODOS vs TODOS, donde los muertos no irán a su mapa de origen, si no que volverán a revivir para tener chances de ganar el evento. El que logre matar más personajes, ganará el evento."

120         Case eModalityEvent.Enfrentamientos

122             If Events(SlotEvent).TeamCant = 1 Then
124                 strDescEvent = "» Los usuarios combatirán en duelos 1vs1"
                Else
126                 strDescEvent = "» Los usuarios combatirán en duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & " donde se escogerán las parejas al azar."
                End If

        End Select

        '<EhFooter>
        Exit Function

strDescEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.strDescEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Sub InitEvent(ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo InitEvent_Err
        '</EhHeader>
    
100     Events(SlotEvent).Run = True
   
102     Select Case Events(SlotEvent).Modality

            Case eModalityEvent.CastleMode
104             Call InitCastleMode(SlotEvent)
                  
106         Case eModalityEvent.DagaRusa
108             Call InitDagaRusa(SlotEvent)
                  
110         Case eModalityEvent.DeathMatch
112             Call InitDeathMatch(SlotEvent)

114         Case eModalityEvent.Busqueda
116             Call InitBusqueda(SlotEvent)
                  
118         Case eModalityEvent.Unstoppable
120             InitUnstoppable SlotEvent
              
122         Case eModalityEvent.Enfrentamientos
                  If Events(SlotEvent).Inscribed <= Events(SlotEvent).QuotasMin Then
                        Events(SlotEvent).LimitRound = Events(SlotEvent).LimitRoundFinal
                  End If
                  
124             Fight_Combate SlotEvent
            
126         Case eModalityEvent.Teleports
128             Call Events_Teleports_Init(SlotEvent)
            
130         Case eModalityEvent.GranBestia
132             'Call Events_GranBestia_Init(SlotEvent)
            
134         Case eModalityEvent.JuegosDelHambre
136             Call Events_JDH_Init(SlotEvent)
                
138         Case Else

                Exit Sub
              
        End Select

        Exit Sub

error:
140     LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : InitEvent() EN EL EVENTO " & Events(SlotEvent).Modality & "."
        '<EhFooter>
        Exit Sub

InitEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.InitEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function CanAttackUserEvent(ByVal UserIndex As Integer, _
                                   ByVal Victima As Integer) As Boolean
        '<EhHeader>
        On Error GoTo CanAttackUserEvent_Err
        '</EhHeader>
          
        ' Si el personaje es del mismo team, no se puede atacar al usuario.
        Dim VictimaSlotUserEvent As Byte
          
100     VictimaSlotUserEvent = UserList(Victima).flags.SlotUserEvent
          
102     If UserList(UserIndex).flags.SlotEvent > 0 And UserList(Victima).flags.SlotEvent > 0 Then

104         With UserList(UserIndex)
                  If Events(.flags.SlotEvent).config(eConfigEvent.eFuegoAmigo) = 0 Then
106                 If Events(.flags.SlotEvent).Users(VictimaSlotUserEvent).Team = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
108                     CanAttackUserEvent = False
    
                        Exit Function
    
                    End If
                End If
                
            End With

        End If
   
110     CanAttackUserEvent = True
          
        Exit Function

error:
112     LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CanAttackUserEvent()"
        '<EhFooter>
        Exit Function

CanAttackUserEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.CanAttackUserEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function PrizeUser_All(ByVal SlotEvent As Byte, _
                               Optional ByVal Team As Byte = 0) As String
        '<EhHeader>
        On Error GoTo PrizeUser_All_Err
        '</EhHeader>
        Dim A As Long
        Dim Temp As String, Text As String
        Dim Bonus As Single
        
100     With Events(SlotEvent)
102         For A = LBound(.Users) To UBound(.Users)
104             If .Users(A).ID > 0 Then
106                 If Team = 0 Or Team = .Users(A).Team Then
                          
108                     Call PrizeUser(.Users(A).ID, Bonus)
                    
110                     Text = Text & UserList(.Users(A).ID).Name & ", "
                    End If
                End If
112         Next A
        
114         If Len(Text) > 2 Then _
                Text = Left$(Text, Len(Text) - 2)
            
116         PrizeUser_All = Text
        End With
        '<EhFooter>
        Exit Function

PrizeUser_All_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.PrizeUser_All " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Private Sub PrizeUser(ByVal UserIndex As Integer, ByVal FinalPoints As Integer)

    On Error GoTo error
          
    ' Premios de los eventos
          
    Dim SlotEvent     As Byte

    Dim SlotUserEvent As Byte

    Dim Obj           As Obj

    SlotEvent = UserList(UserIndex).flags.SlotEvent
    SlotUserEvent = UserList(UserIndex).flags.SlotUserEvent
          
    With Events(SlotEvent)
             
        If .PrizeGld = 0 And .PrizeEldhir = 0 And .PrizePoints = 0 And .PrizeObj.ObjIndex = 0 Then Exit Sub
   
        If .PrizeGld > 0 Then
            UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + .PrizeGld
            Call WriteUpdateGold(UserIndex)
        End If
                    
        If .PrizePoints > 0 Then
            If FinalPoints > 0 Then
                UserList(UserIndex).Stats.Points = UserList(UserIndex).Stats.Points + FinalPoints
            Else
                UserList(UserIndex).Stats.Points = UserList(UserIndex).Stats.Points + .PrizePoints
            End If
        End If
                    
        If .PrizeObj.ObjIndex > 0 Then
            If Not MeterItemEnInventario(UserIndex, .PrizeObj) Then
                WriteConsoleMsg UserIndex, "Tu premio OBJETO no ha sido entregado, envia esta foto a un Game Master.", FontTypeNames.FONTTYPE_INFO
                LogEventos ("Personaje " & UserList(UserIndex).Name & " no recibió: " & .PrizeObj.ObjIndex & " (x" & .PrizeObj.Amount & ")")
            End If

        End If

              
        Call Events_Reward_User(UserIndex, SlotEvent)

        'UserList(UserIndex).Stats.TorneosGanados = UserList(UserIndex).Stats.TorneosGanados + 1

    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : PrizeUser()"
End Sub

Private Sub ChangeBodyEvent(ByVal UserIndex As Integer, _
                            ByVal ChangeHead As Boolean)

    On Error GoTo error

    ' En caso de que el evento cambie el body, de lo cambiamos.
    With UserList(UserIndex)
        ' Si ya está mimetizado tenemos que cambiar su apariencia y no guardar nada
        If Not .flags.Mimetizado > 0 Then
            .flags.Mimetizado = 1
            .CharMimetizado.Body = .Char.Body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        End If
        
        .Char.Body = Events_ChangeBody()
        .Char.Head = IIf(ChangeHead = False, .Char.Head, 0)
        .Char.CascoAnim = 0
        .Char.ShieldAnim = 0
        .Char.WeaponAnim = 0
        
        
        ChangeUserChar UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex
        RefreshCharStatus UserIndex
          
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : ChangeBodyEvent()"
End Sub

Private Function ResetBodyEvent(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

    On Error GoTo error

    ' En caso de que el evento cambie el body del personaje, se lo restauramos.
          
    With UserList(UserIndex)

        If .flags.Muerto Then
            Call RevivirUsuario(UserIndex)
            Exit Function
        End If
        
        'If Events(SlotEvent).Users(.flags.SlotUserEvent).Selected = 0 Then Exit Function
              
        If .CharMimetizado.Body > 0 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                  
            .CharMimetizado.Body = 0
            .CharMimetizado.Head = 0
            .CharMimetizado.CascoAnim = 0
            .CharMimetizado.ShieldAnim = 0
            .CharMimetizado.WeaponAnim = 0
            
            Dim A As Long
            
            For A = 1 To 4
                .CharMimetizado.AuraIndex(A) = 0
            Next A
            
            .ShowName = True
                  
            ChangeUserChar UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex
            RefreshCharStatus UserIndex
        End If
          
    End With
          
    Exit Function

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : ResetBodyEvent()"
End Function
Public Function Events_GenerateSpamShort(ByVal Slot As Byte) As String
    Dim strTemp As String

    With Events(Slot)

        Events_GenerateSpamShort = "Evento '" & UCase$(.Name) & IIf(.config(eConfigEvent.eParty) = 1, " (Party)", "") & ". " & strTemp & _
            "Ganá hasta " & IIf(.PrizePoints > 0, .PrizePoints & " Puntos de Torneo", "0 Puntos") & ". Únete con '/TORNEO " & UCase$(.Name) & "'." & vbCrLf & "El evento comienza en " & SecondsToHMS(.TimeCancel)
               
    End With
End Function
Public Function Events_GenerateSpam(ByVal Slot As Byte) As String

    Dim strTemp  As String

    Dim txtRojas As String
    
    With Events(Slot)

        ' Event introduction
        strTemp = "_______________________________________"
        ' Modality
        strTemp = strTemp & vbCrLf & "'" & UCase$(.Name) & IIf(.config(eConfigEvent.eFuegoAmigo) = 1, " (Fuego Amigo)", vbNullString)

        ' Rounds
        If (.Modality = Enfrentamientos) Then
            strTemp = strTemp & " | Rounds: " & .LimitRound & IIf(.LimitRound > 1, "s", vbNullString) & IIf(.LimitRoundFinal <> .LimitRound, ". (Final a " & .LimitRoundFinal & ")", vbNullString)
        End If

        ' Points prize
        If .PrizePoints > 0 Then
            strTemp = strTemp & vbCrLf & "¡El Pozo Mayor es " & .PrizePoints & " Puntos de Torneo! "
        End If


        ' Special rules
        If .Modality = eModalityEvent.DagaRusa Then
            strTemp = strTemp & vbCrLf & "Nota: El Asesino tiene " & .Prob & "% de probabilidad de apuñalarte."
        End If
        

        ' Level requirement
        If Not (.LvlMin = 1 And .LvlMax = 47) And Not (.LvlMin = 1 And .LvlMax = 1) Then
            strTemp = strTemp & vbCrLf & "Requisito de nivel: " & .LvlMin & " a " & .LvlMax & ". "
        End If

        ' Class and faction requirement
        Dim TextClass As String: TextClass = CheckAllowedClasses(.AllowedClasses)
        Dim TextFaction As String: TextFaction = Events_CheckAllowed_Faction(.AllowedFaction)
        If TextClass <> "TODAS" Or TextFaction <> "TODAS" Then
            strTemp = strTemp & vbCrLf & IIf(TextClass <> "TODAS", "Clases permitidas - " & TextClass, "") & IIf(TextFaction <> "TODAS", " | Facciones permitidas - " & TextFaction, "")
        End If

        ' Fees
        If .InscriptionGld > 0 Or .InscriptionEldhir > 0 Then
            strTemp = strTemp & vbCrLf & "Valor Inscripción (ORO): " & IIf(.InscriptionGld > 0, .InscriptionGld & " de oro. (Llevate el pozo acumulado) ", "") & IIf(.InscriptionEldhir > 0, " | " & .InscriptionEldhir & " GldPremiums.", "")
        End If

        ' Prizes
        If .PrizeGld > 0 Or .PrizeEldhir > 0 Or .PrizeObj.ObjIndex > 0 Then
            strTemp = strTemp & vbCrLf & "Premios: " & IIf(.PrizeGld > 0, .PrizeGld & " de oro", "") & IIf(.PrizeEldhir > 0, " | " & .PrizeEldhir & " DSP", "")
            If .PrizeObj.Amount > 0 Then
                strTemp = strTemp & IIf(.PrizeObj.ObjIndex > 0, " | Objeto: " & ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ")", "")
            End If
        Else
            If .PrizePoints = 0 Then
            strTemp = strTemp & vbCrLf & "Nota: Evento sin premios."
            End If
        End If


        ' Special rules continued
        If .config(eConfigEvent.eCascoEscudo) = 0 Then
            strTemp = strTemp & vbCrLf & "Regla especial: No se permiten Cascos-Escudos. "
        End If

        ' Items restrictions
        If .config(eConfigEvent.eBronce) = 0 Or .config(eConfigEvent.ePlata) = 0 Or .config(eConfigEvent.eOro) = 0 Or .config(eConfigEvent.ePremium) = 0 Then
            strTemp = strTemp & vbCrLf & "Restricciones de items: "
            If .config(eConfigEvent.eBronce) = 0 Then strTemp = strTemp & " [BRONCE]"
            If .config(eConfigEvent.ePlata) = 0 Then strTemp = strTemp & " [PLATA]"
            If .config(eConfigEvent.eOro) = 0 Then strTemp = strTemp & " [ORO]"
            If .config(eConfigEvent.ePremium) = 0 Then strTemp = strTemp & " [PREMIUM]"
        End If

        ' Special spells
        If .config(eConfigEvent.eResu) = 1 Or .config(eConfigEvent.eInvisibilidad) = 1 Or .config(eConfigEvent.eOcultar) = 1 Or .config(eConfigEvent.eInvocar) = 1 Then
            strTemp = strTemp & vbCrLf & "Hechizos NO permitidos: "
            If .config(eConfigEvent.eResu) = 1 Then strTemp = strTemp & " 'RESU' "
            If .config(eConfigEvent.eInvisibilidad) = 1 Then strTemp = strTemp & " 'INVI' "
            If .config(eConfigEvent.eOcultar) = 1 Then strTemp = strTemp & " 'OCULTAR' "
            If .config(eConfigEvent.eInvocar) = 1 Then strTemp = strTemp & " 'INVOCAR'"
        End If

        ' How to join
        strTemp = strTemp & vbCrLf & "Para inscribirte tipea el comando: '/TORNEO " & UCase$(.Name) & "'"
        
        strTemp = strTemp & vbCrLf & "_______________________________________"
    End With

    Events_GenerateSpam = strTemp

End Function
Public Function Events_GenerateSpamDiscord(ByVal Slot As Byte) As String

    Dim strTemp  As String

    Dim txtRojas As String
    
    With Events(Slot)

        ' Event introduction
        strTemp = "__"
        ' Modality
        strTemp = strTemp & vbCrLf & "**'" & UCase$(.Name) & IIf(.config(eConfigEvent.eFuegoAmigo) = 1, "'** (Fuego Amigo)", vbNullString)

        ' Rounds
        If (.Modality = Enfrentamientos) Then
            strTemp = strTemp & " | Rounds: " & .LimitRound & IIf(.LimitRound > 1, "s", vbNullString) & IIf(.LimitRoundFinal <> .LimitRound, ". **(Final a " & .LimitRoundFinal & ")**", vbNullString)
        End If
        
        ' Special rules
        If .Modality = eModalityEvent.DagaRusa Then
            strTemp = strTemp & vbCrLf & "**Nota:** El Asesino tiene **" & .Prob & "%** de probabilidad de apuñalarte."
        End If
        

        ' Level requirement
        If Not (.LvlMin = 1 And .LvlMax = 47) And Not (.LvlMin = 1 And .LvlMax = 1) Then
            strTemp = strTemp & vbCrLf & "Requisito de nivel: " & .LvlMin & " a " & .LvlMax & ". "
        End If

        ' Class and faction requirement
        Dim TextClass As String: TextClass = CheckAllowedClasses(.AllowedClasses)
        Dim TextFaction As String: TextFaction = Events_CheckAllowed_Faction(.AllowedFaction)
        If TextClass <> "TODAS" Or TextFaction <> "TODAS" Then
            strTemp = strTemp & vbCrLf & IIf(TextClass <> "TODAS", "Clases permitidas - **" & TextClass & "**", "") & IIf(TextFaction <> "TODAS", " | Facciones permitidas - " & TextFaction, "")
        End If

        ' Fees
        If .InscriptionGld > 0 Or .InscriptionEldhir > 0 Then
            strTemp = strTemp & vbCrLf & "Valor **Inscripción (ORO)**: " & IIf(.InscriptionGld > 0, "**" & .InscriptionGld & "** de oro. (**Llevate el** poso **acumulado**) ", "") & IIf(.InscriptionEldhir > 0, " | " & .InscriptionEldhir & " GldPremiums.", "")
        End If

        ' Points prize
        If .PrizePoints > 0 Then
            strTemp = strTemp & vbCrLf & "¡Gana HASTA " & .PrizePoints & " **Puntos de Torneo**! "
        End If
        
        ' Prizes
        If .PrizeGld > 0 Or .PrizeEldhir > 0 Or .PrizeObj.ObjIndex > 0 Then
            strTemp = strTemp & vbCrLf & "**Premios:** " & IIf(.PrizeGld > 0, .PrizeGld & " de **ORO**", "") & IIf(.PrizeEldhir > 0, " | " & .PrizeEldhir & " **DSP**", "")
            If .PrizeObj.Amount > 0 Then
                strTemp = strTemp & IIf(.PrizeObj.ObjIndex > 0, " | **Objeto: " & ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ")**", "")
            End If
        Else
            If .PrizePoints = 0 Then
            strTemp = strTemp & vbCrLf & "**Nota:** Evento sin premios."
            End If
        End If


        ' Special rules continued
        If .config(eConfigEvent.eCascoEscudo) = 0 Then
            strTemp = strTemp & vbCrLf & "**Regla especial:** No se permiten Cascos-Escudos. "
        End If

        ' Items restrictions
        If .config(eConfigEvent.eBronce) = 0 Or .config(eConfigEvent.ePlata) = 0 Or .config(eConfigEvent.eOro) = 0 Or .config(eConfigEvent.ePremium) = 0 Then
            strTemp = strTemp & vbCrLf & "**Restricciones** de items: "
            If .config(eConfigEvent.eBronce) = 0 Then strTemp = strTemp & " [BRONCE]"
            If .config(eConfigEvent.ePlata) = 0 Then strTemp = strTemp & " [PLATA]"
            If .config(eConfigEvent.eOro) = 0 Then strTemp = strTemp & " [ORO]"
            If .config(eConfigEvent.ePremium) = 0 Then strTemp = strTemp & " [PREMIUM]"
        End If

        ' Special spells
        If .config(eConfigEvent.eResu) = 1 Or .config(eConfigEvent.eInvisibilidad) = 1 Or .config(eConfigEvent.eOcultar) = 1 Or .config(eConfigEvent.eInvocar) = 1 Then
            strTemp = strTemp & vbCrLf & "Hechizos **NO** permitidos: "
            If .config(eConfigEvent.eResu) = 1 Then strTemp = strTemp & " **'RESU'** "
            If .config(eConfigEvent.eInvisibilidad) = 1 Then strTemp = strTemp & " **'INVI'** "
            If .config(eConfigEvent.eOcultar) = 1 Then strTemp = strTemp & " **'OCULTAR'** "
            If .config(eConfigEvent.eInvocar) = 1 Then strTemp = strTemp & " **'INVOCAR'**"
        End If

        ' How to join
        strTemp = strTemp & vbCrLf & "Para **inscribirte** tipea el comando: **'/TORNEO " & UCase$(.Name) & "'**"
        strTemp = strTemp & vbCrLf & "Tienes **15 minutos** antes de que se cancele."
        strTemp = strTemp & vbCrLf & "__"
    End With

    Events_GenerateSpamDiscord = strTemp

End Function

Private Function Events_CheckNew(ByRef Data As tEvents) As Boolean
        '<EhHeader>
        On Error GoTo Events_CheckNew_Err
        '</EhHeader>

100     Events_CheckNew = False
102     If Data.PrizeObj.ObjIndex > NumObjDatas Then Exit Function
104     If Data.PrizeObj.Amount < 0 Or Data.PrizeObj.Amount > 1000 Then Exit Function
106     Events_CheckNew = True
    
        '<EhFooter>
        Exit Function

Events_CheckNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_CheckNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function NewEvent(ByRef Data As tEvents, _
                         Optional ByVal Predeterminado As Boolean, _
                         Optional NickGm As String = vbNullString) As Byte
                          
    On Error GoTo error
                          
    Dim Slot    As Integer

    Dim strTemp As String
    
    Slot = FreeSlotEvent()
    
    If Not Events_CheckNew(Data) Then
        NewEvent = 0
        Exit Function
    End If
    
    If Event_NameRepeat(UCase$(Data.Name)) Then
        NewEvent = 0
        Exit Function
    End If
        
    If Slot = 0 Then
        NewEvent = 0
        Exit Function
    Else
        NewEvent = Slot
        Events(Slot) = Data
        ReDim Events(Slot).Users(1 To Events(Slot).QuotasMax) As tUserEvent
              
        
        If Not Predeterminado Then
            Dim TextSpam As String
            Dim TextSpamDiscord As String
            
            TextSpam = Events_GenerateSpam(Slot)
            TextSpamDiscord = Events_GenerateSpamDiscord(Slot)
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(TextSpam, FontTypeNames.FONTTYPE_EVENT, cEvents_General)
            
           If NickGm <> vbNullString Then
                Call LogEventos("NICK GM " & NickGm & vbCrLf & TextSpam & vbCrLf)
                
                ' # Envia un mensaje al canal de DISCORD de TORNEOS
                WriteMessageDiscord CHANNEL_TOURNAMENT, TextSpamDiscord
           End If
        End If
        
        EventLast = EventLast + 1
    End If

    Exit Function

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : NewEvent()"
End Function

Private Sub GiveBack_Inscription(ByVal SlotEvent As Byte)

    On Error GoTo error

    Dim LoopC As Integer

    Dim Obj   As Obj
          
    With Events(SlotEvent)
          
        Obj.ObjIndex = 880
        Obj.Amount = .InscriptionEldhir
              
        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                If .InscriptionEldhir > 0 Then
                    UserList(.Users(LoopC).ID).Stats.Eldhir = UserList(.Users(LoopC).ID).Stats.Eldhir + .InscriptionEldhir
                    WriteUpdateDsp (.Users(LoopC).ID)
                End If
                      
                If .InscriptionGld > 0 Then
                    UserList(.Users(LoopC).ID).Stats.Gld = UserList(.Users(LoopC).ID).Stats.Gld + .InscriptionGld
                    WriteUpdateGold (.Users(LoopC).ID)
                End If
            End If

        Next LoopC

    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : GiveBack_Inscription()"
End Sub

Public Sub CloseEvent(ByVal Slot As Byte, _
                      Optional ByVal MsgConsole As String = vbNullString, _
                      Optional ByVal Cancel As Boolean = False)

    On Error GoTo error
          
    With Events(Slot)

        ' Devolvemos la inscripción
        If Cancel Then
            Call GiveBack_Inscription(Slot)
        End If
              
        If MsgConsole <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(MsgConsole, FontTypeNames.FONTTYPE_EVENT, eMessageType.cEvents_Curso)
              
        Call ResetEvent(Slot)
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CloseEvent()"
End Sub

Public Function Events_Group_CheckUsers(ByVal GroupIndex As Integer, ByVal SlotEvent As Byte, ByRef ErrorMsg As String) As Boolean

    Dim A As Long
    Dim tUser As Integer
    
    For A = 1 To MAX_MEMBERS_GROUP
        tUser = Groups(GroupIndex).User(A).Index
            
        If tUser > 0 Then
            If Not Events_CheckUserEvent(tUser, SlotEvent, ErrorMsg) Then
                Exit Function
            End If
        End If
    Next A
    
    Events_Group_CheckUsers = True
    
End Function
Public Sub Events_Group_Set(ByVal GroupIndex As Integer, ByVal SlotEvent As Byte)
    
    Dim A As Long
    Dim tUser As Integer
    
    For A = 1 To MAX_MEMBERS_GROUP
        tUser = Groups(GroupIndex).User(A).Index
                
        If tUser > 0 Then
            Call Event_SetNewUser(tUser, SlotEvent)
        End If
    Next A
End Sub


' # El personaje es seteado dentro del evento.
Public Sub Event_SetNewUser(ByVal UserIndex As Integer, ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo Event_SetNewUser_Err
        '</EhHeader>

        Dim Pos As WorldPos
        Dim Slot As Byte
        
100     Slot = FreeSlotUser(SlotEvent)
        
        Events(SlotEvent).Inscribed = Events(SlotEvent).Inscribed + 1
        Events(SlotEvent).Users(Slot).Name = UCase$(UserList(UserIndex).Name)
        
        Call WriteConsoleMsg(UserIndex, "¡Has sido anotado para la partida " & Events(SlotEvent).Name & "!", FontTypeNames.FONTTYPE_INFOGREEN)
        
        
        Dim Message As String
        Message = "Partida " & Events(SlotEvent).Name & "» Se anota " & UserList(UserIndex).Name & _
                    " [" & ListaClases(UserList(UserIndex).Clase) & " " & ListaRazas(UserList(UserIndex).Raza) & "] Inscriptos TOTALES: " & Events(SlotEvent).Inscribed
        
        Call SendData(SendTarget.toMap, 1, PrepareMessageRenderConsole(Message, d_AddGld, 6000, 0))
        
        '<EhFooter>
        Exit Sub

Event_SetNewUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_SetNewUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Busca un usuario en el evento al que quiere participar.
Public Function Event_CheckExistUser(ByVal UserIndex As Integer, ByVal SlotEvent As Byte) As Integer
        '<EhHeader>
        On Error GoTo Event_SetNewUser_Err
        '</EhHeader>

        Dim A As Long
        Dim uName As String
        
        uName = UCase$(UserList(UserIndex).Name)
        
        With Events(SlotEvent)
            For A = LBound(.Users) To UBound(.Users)
                If StrComp(.Users(A).Name, uName) = 0 Then
                    Event_CheckExistUser = A
                    Exit Function
                End If
            Next A
        
        End With
        
        
        
        '<EhFooter>
        Exit Function

Event_SetNewUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_CheckExistUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' # El usuario es teletransportado para participar de otro evento.
Public Sub ParticipeEvent_User(ByVal UserIndex As Integer, ByVal SlotEvent As Byte)

    Dim Pos As WorldPos
    
    With UserList(UserIndex)
        .flags.SlotEvent = SlotEvent
        .flags.SlotUserEvent = Event_CheckExistUser(UserIndex, SlotEvent)
        
        Events(SlotEvent).Users(.flags.SlotUserEvent).ID = UserIndex
        
        .Stats.Gld = .Stats.Gld - Events(SlotEvent).InscriptionGld
        .Stats.Eldhir = .Stats.Eldhir - Events(SlotEvent).InscriptionEldhir
        
110         .PosAnt.Map = .Pos.Map
112         .PosAnt.X = .Pos.X
114         .PosAnt.Y = .Pos.Y
                  
        
122         Pos.Map = MapEvent.SalaEspera.Map
124         Pos.X = MapEvent.SalaEspera.X
126         Pos.Y = MapEvent.SalaEspera.Y
                      
128     Call FindLegalPos(UserIndex, Pos.Map, Pos.X, Pos.Y)
130     Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
        Call WriteUpdateUserStats(UserIndex)
    End With
End Sub

Public Function Event_CheckInscriptions_User(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByRef ErrorMsg As String) As Boolean


    With UserList(UserIndex)
        '# Comprueba si tiene grupo formado. En ese caso chequea todo lo necesario para inscribirte con él. Sino intentará unirte Azar.
        If .GroupIndex > 0 Then
            If Groups(.GroupIndex).Members <> Events(SlotEvent).TeamCant Then
                ErrorMsg = "El grupo debe tener " & Events(SlotEvent).TeamCant & " miembros para participar con un equipo. Sino puedes optar por entrar solo y encontrar pareja al Azar."
                Exit Function
            End If
            
            If Not Events_Group_CheckUsers(.GroupIndex, SlotEvent, ErrorMsg) Then
                ' Message Internal
                Exit Function
            End If
        
        Else
            If Not Events_CheckUserEvent(UserIndex, SlotEvent, ErrorMsg) Then
                ' Message Internal
                Exit Function
            End If
        End If
    
    
    Event_CheckInscriptions_User = True
    End With
End Function

Public Sub Event_Initial(ByVal SlotEvent As Byte, ByVal Inscribed As Byte)
        '<EhHeader>
        On Error GoTo Event_Initial_Err
        '</EhHeader>

100     With Events(SlotEvent)
102         .Quotas = Inscribed
104         .TimeCancel = 0
                
122         LogEventos "Cupos alcanzados."
126         If .PrizePoints > 0 Then .PrizePoints = Events_SetReward_Points(Events(SlotEvent), .Quotas)
            If .InscriptionGld > 0 Then .InscriptionGld = Events_SetInscription_Gold(Events(SlotEvent), .Quotas)
            
            .TimeInit = 5
        End With

        '<EhFooter>
        Exit Sub

Event_Initial_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_Initial " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Event_ClassOld(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Event_ClassOld_Err
        '</EhHeader>

100     With Events(SlotEvent)

            Dim A As Long

102         With UserList(UserIndex)

104             If .flags.TomoPocion Then
106                 LogEventos ("El personaje " & .Name & " era: " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & ". Atributos: " & .Stats.UserAtributosBackUP(1) & "-" & .Stats.UserAtributosBackUP(2) & "-" & .Stats.UserAtributosBackUP(3) & "-" & .Stats.UserAtributosBackUP(4) & "-" & .Stats.UserAtributosBackUP(5) & ". Oro: " & .Stats.Gld & ", D.Azules: " & .Stats.Eldhir & ". Vida: " & .Stats.MaxHp & " y maná " & .Stats.MaxMan & ". HIT: " & .Stats.MaxHit)
                Else
108                 LogEventos ("El personaje " & .Name & " era: " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & ". Atributos: " & .Stats.UserAtributos(1) & "-" & .Stats.UserAtributos(2) & "-" & .Stats.UserAtributos(3) & "-" & .Stats.UserAtributos(4) & "-" & .Stats.UserAtributos(5) & ". Oro: " & .Stats.Gld & ", D.Azules: " & .Stats.Eldhir & ". Vida: " & .Stats.MaxHp & " y maná " & .Stats.MaxMan & ". HIT: " & .Stats.MaxHit)
                End If

                Dim OldChar      As String
        
                Dim FilePath_Old As String
                Dim FilePath_Copy As String
                
                FilePath_Old = CharPath & UCase$(UserList(UserIndex).Name) & ".chr"
                
                Call SaveUser(UserList(UserIndex), FilePath_Old)
                FilePath_Copy = Replace(CharPath, "CHARS\CHARFILE", "CHARS\CHARFILE_EVENTS_COPY") & UCase$(UserList(UserIndex).Name) & ".chr"
                
                Call FileCopy(FilePath_Old, FilePath_Copy)
            End With
    
        End With

        '<EhFooter>
        Exit Sub

Event_ClassOld_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_ClassOld " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AbandonateEvent(ByVal UserIndex As Integer, _
                           Optional ByVal MsgAbandonate As Boolean = False, _
                           Optional ByVal Forzado As Boolean = False, _
                           Optional ByVal ResetTotal As Boolean = False)


On Error GoTo ErrHandler
    Dim Pos           As WorldPos

    Dim SlotEvent     As Byte

    Dim SlotUserEvent As Byte

    Dim UserTeam      As Byte

    Dim UserMapFight  As Byte
          
    With UserList(UserIndex)
        SlotEvent = .flags.SlotEvent
        SlotUserEvent = .flags.SlotUserEvent
              
        
        If SlotEvent > 0 And SlotUserEvent > 0 Then

            With Events(SlotEvent)
                .Inscribed = .Inscribed - 1
                UserTeam = .Users(SlotUserEvent).Team
                UserMapFight = .Users(SlotUserEvent).MapFight
     
                Select Case .Modality

                    Case eModalityEvent.JuegosDelHambre

                        If Forzado Then
                            Call TirarTodosLosItems(UserIndex)

                        End If
                        
                    Case eModalityEvent.DagaRusa

                        If .Run Then

                            Call WriteUserInEvent(UserIndex)
                        End If
                                  
                    Case eModalityEvent.Enfrentamientos
                    
                        If Not Forzado Then
                            If UserList(UserIndex).Counters.TimeFight > 0 Then
                                UserList(UserIndex).Counters.TimeFight = 0
                                Call WriteUserInEvent(UserIndex)

                            End If

                        End If
                        
                    Case eModalityEvent.DeathMatch
                        UserList(UserIndex).flags.Mimetizado = 0
                                  
                End Select
                          
                If UserList(UserIndex).Counters.TimeFight > 0 Then
                    Call WriteRender_CountDown(UserIndex, 0)
                End If
                
                ' Pociones Rojas que le hayan quedado
                If .Modality = Enfrentamientos Then
                    Call QuitarObjetos(POCION_ROJA, MAX_INVENTORY_OBJS, UserIndex)
                End If
                
                Pos.Map = UserList(UserIndex).PosAnt.Map
                Pos.X = UserList(UserIndex).PosAnt.X
                Pos.Y = UserList(UserIndex).PosAnt.Y
                          
                Call FindLegalPos(UserIndex, Pos.Map, Pos.X, Pos.Y)
                Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
                          
                If Events(SlotEvent).config(eConfigEvent.eMezclarApariencias) = 1 Then
                                                    
                    Call ResetBodyEvent(SlotEvent, UserIndex)

                End If
                  
                UserList(UserIndex).ShowName = True
                
                          
                If MsgAbandonate Then WriteConsoleMsg UserIndex, "Has abandonado el evento. Podrás recibir una pena por hacer esto.", FontTypeNames.FONTTYPE_WARNING
                
                .Users(SlotUserEvent).ID = 0
                .Users(SlotUserEvent).Team = 0
                
                If .Run Then
                    'UserList(UserIndex).Stats.TorneosJugados = UserList(UserIndex).Stats.TorneosJugados + 1
                    .Users(SlotUserEvent).Value = 0
                    .Users(SlotUserEvent).Selected = 0
                    .Users(SlotUserEvent).MapFight = 0
                    .Users(SlotUserEvent).Oponent = 0
                    .Users(SlotUserEvent).RoundsWin = 0
                    .Users(SlotUserEvent).RoundsWinFinal = 0
                    .Users(SlotUserEvent).Damage = 0
                    .Users(SlotUserEvent).TimeCancel = 0
                End If
                
                UserList(UserIndex).flags.SlotEvent = 0
                UserList(UserIndex).flags.SlotUserEvent = 0
                UserList(UserIndex).flags.FightTeam = 0
                UserList(UserIndex).Counters.TimeApparience = 0
                UserList(UserIndex).Counters.TimeTelep = 0
                
                RefreshCharStatus UserIndex
                
                Call WriteUpdateUserStats(UserIndex)
                
                'Call Streamer_CheckUser(UserIndex)
                
                
            End With
            

        End If
        
    End With
    
    Exit Sub
ErrHandler:
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Error", 1))
          
End Sub

' @ Un evento es cerrado tras finalizar su tiempo de ejecución pactado.
Private Sub Events_FinishTime(ByVal SlotEvent As Byte)
    
    On Error GoTo ErrHandler
    
    Dim CopyUsers() As tUserEvent
    Dim UserName As String
    Dim Bonus As Single
    
    With Events(SlotEvent)
        Select Case .Modality
            Case eModalityEvent.Unstoppable, _
                 eModalityEvent.Busqueda
                 
                Events_OrdenateUsersValue SlotEvent, CopyUsers
                
                UserName = UserList(CopyUsers(1).ID).Name
                Call PrizeUser(CopyUsers(1).ID, Bonus)
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» Ganador " & UserName & ". " & Events_StrReward(SlotEvent) & vbCrLf & Event_GenerateTablaPos(SlotEvent, CopyUsers), FontTypeNames.FONTTYPE_EVENT))
        End Select
    
        CloseEvent (SlotEvent)
    End With
    
    Exit Sub
ErrHandler:
End Sub

'#################EVENTO CASTLE MODE##########################
Public Function CanAttackReyCastle(ByVal UserIndex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo CanAttackReyCastle_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If .flags.SlotEvent > 0 Then
104             If Npclist(NpcIndex).flags.TeamEvent = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
106                 CanAttackReyCastle = False

                    Exit Function

                End If
            End If
          
108         CanAttackReyCastle = True
        End With

        '<EhFooter>
        Exit Function

CanAttackReyCastle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.CanAttackReyCastle " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Sub CastleMode_InitRey()

    On Error GoTo error
          
    Dim NpcIndex As Integer

    Const NumRey As Integer = 697

    Dim Pos      As WorldPos

    Dim LoopX    As Integer, LoopY As Integer

    Const Rango  As Byte = 5
          
    For LoopX = YMinMapSize To YMaxMapSize
        For LoopY = XMinMapSize To XMaxMapSize

            If InMapBounds(MapEvent.Castle(0).Map, LoopX, LoopY) Then
                If MapData(MapEvent.Castle(0).Map, LoopX, LoopY).NpcIndex > 0 Then
                    Call QuitarNPC(MapData(MapEvent.Castle(0).Map, LoopX, LoopY).NpcIndex)
                End If
            End If

        Next LoopY
    Next LoopX
          
    Pos.Map = MapEvent.Castle(0).Map
                  
    Pos.X = MapEvent.Castle(0).X
    Pos.Y = MapEvent.Castle(0).Y
    NpcIndex = SpawnNpc(NumRey, Pos, False, False)
    Npclist(NpcIndex).flags.TeamEvent = 2
              
    Pos.X = MapEvent.Castle(1).X
    Pos.Y = MapEvent.Castle(1).Y
    NpcIndex = SpawnNpc(NumRey, Pos, False, False)
    Npclist(NpcIndex).flags.TeamEvent = 1
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CastleMode_InitRey()"
End Sub

Public Sub InitCastleMode(ByVal SlotEvent As Byte)

    On Error GoTo error

    Dim LoopC    As Integer

    Dim NpcIndex As Integer

    Dim Pos      As WorldPos
          
    ' Spawn the npc castle mode
    CastleMode_InitRey
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                If LoopC > (UBound(.Users()) / 2) Then
                    .Users(LoopC).Team = 2
                    UserList(.Users(LoopC).ID).flags.FightTeam = 2
                    Pos.Map = MapEvent.Castle(0).Map
                    Pos.X = MapEvent.Castle(0).X
                    Pos.Y = MapEvent.Castle(0).Y
                          
                    Call FindLegalPos(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y)
                    Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, False)
                Else
                    .Users(LoopC).Team = 1
                    UserList(.Users(LoopC).ID).flags.FightTeam = 1
                    Pos.Map = MapEvent.Castle(1).Map
                    Pos.X = MapEvent.Castle(1).X
                    Pos.Y = MapEvent.Castle(1).Y
                          
                    Call FindLegalPos(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y)
                    Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, False)
                          
                End If
            End If

        Next LoopC

    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : InitCastleMode()"
End Sub

Public Sub FinishCastleMode(ByVal SlotEvent As Byte, ByVal UserEventSlot As Integer)

    On Error GoTo error

    Dim LoopC     As Integer

    Dim strTemp   As String

    Dim NpcIndex  As Integer

    Dim MiObj     As Obj
          
    With Events(SlotEvent)
        MiObj.ObjIndex = 899
        MiObj.Amount = 1
                                  
        If Not MeterItemEnInventario(.Users(UserEventSlot).ID, MiObj) Then
            Call LogEventos("Recompensa del Rey no entregada")
        End If
                                  
        MiObj.ObjIndex = 900
        MiObj.Amount = 1
                                  
        If Not MeterItemEnInventario(.Users(UserEventSlot).ID, MiObj) Then
            Call LogEventos("Recompensa del Rey no entregada")
        End If
        
        Call WriteConsoleMsg(.Users(UserEventSlot).ID, "Aquí tienes algo del Rey enemigo ¡Que lo disfrutes!", FontTypeNames.FONTTYPE_ANGEL)
        
        For LoopC = LBound(.Users()) To UBound(.Users())
            
            If .Users(LoopC).ID > 0 Then
                If .Users(LoopC).Team = .Users(UserEventSlot).Team Then
                    Call PrizeUser(.Users(LoopC).ID, 0)
                          
                    If strTemp = vbNullString Then
                        strTemp = UserList(.Users(LoopC).ID).Name
                    Else
                        strTemp = strTemp & ", " & UserList(.Users(LoopC).ID).Name
                    End If
                End If
            End If

        Next LoopC
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» Ganadores " & strTemp & ". " & Events_StrReward(SlotEvent), FontTypeNames.FONTTYPE_EVENT))
        Call CloseEvent(SlotEvent)
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : FinishCastleMode()"
End Sub

' FIN EVENTO CASTLE MODE #####################################

' ###################### EVENTO DAGA RUSA ###########################
Public Sub InitDagaRusa(ByVal SlotEvent As Byte)

    On Error GoTo error

    Dim LoopC    As Integer

    Dim NpcIndex As Integer

    Dim Pos      As WorldPos
    
    Dim Y As Long
    
    
    With Events(SlotEvent)

        Call DataRusa_SummonUser(SlotEvent, True)
              
        Pos.Map = MapEvent.DagaRusa.Map
        Pos.X = MapEvent.DagaRusa.X
        Pos.Y = MapEvent.DagaRusa.Y - 1
            
        NpcIndex = SpawnNpc(316, Pos, True, False)
          
        If NpcIndex <> 0 Then
            Npclist(NpcIndex).Movement = NpcDagaRusa
            Npclist(NpcIndex).Hostile = 0
            
            Npclist(NpcIndex).flags.SlotEvent = SlotEvent
            Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed
            .NpcIndex = NpcIndex
                  
            Events_AI_DagaRusa NpcIndex, True
        End If
              
        .TimeCount = 4
    End With

    Exit Sub

error:
    LogError "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : InitDagaRusa()"
End Sub

' @ Sumonea a todos los que están vigentes en la DAGA
Public Sub DataRusa_SummonUser(ByVal SlotEvent As Byte, Optional ByVal Blocked As Boolean)
        '<EhHeader>
        On Error GoTo DataRusa_SummonUser_Err
        '</EhHeader>

        Dim A As Integer
        Dim Num As Long
        Dim posY As Integer
        
100     With Events(SlotEvent)

102         For A = LBound(.Users()) To UBound(.Users())

104             If .Users(A).ID > 0 Then
106                 If Blocked Then
108                     Call WriteUserInEvent(.Users(A).ID)
                    End If
                    
                    If EsPar(A) Then
                        posY = MapEvent.DagaRusa.Y - 2
                        Num = Num + 1
                    Else
                        posY = MapEvent.DagaRusa.Y
                    End If
                    
110                 Call WarpUserChar(.Users(A).ID, MapEvent.DagaRusa.Map, MapEvent.DagaRusa.X + Num, MapEvent.DagaRusa.Y, False)
                    
                End If

112         Next A
    
        End With

        '<EhFooter>
        Exit Sub

DataRusa_SummonUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.DataRusa_SummonUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function DagaRusa_NextUser(ByVal SlotEvent As Byte) As Byte
        '<EhHeader>
        On Error GoTo DagaRusa_NextUser_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     DagaRusa_NextUser = 0
          
102     With Events(SlotEvent)

104         For LoopC = LBound(.Users()) To UBound(.Users())

106             If (.Users(LoopC).ID > 0) And (.Users(LoopC).Value = 0) Then
108                 DagaRusa_NextUser = .Users(LoopC).ID

                    '.Users(LoopC).Value = 1
                    Exit For

                End If

110         Next LoopC

        End With

        '<EhFooter>
        Exit Function

DagaRusa_NextUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.DagaRusa_NextUser " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub DagaRusa_ResetRonda(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())
            .Users(LoopC).Value = 0
        Next LoopC
          
    End With

End Sub

' # Se finaliza el evento
Private Sub Events_Finish(ByVal SlotEvent As Byte, ByVal TeamUser As Byte)
        '<EhHeader>
        On Error GoTo Events_Finish_Err
        '</EhHeader>
    
        Dim TextWinner As String
    
100     With Events(SlotEvent)
            TextWinner = Events_Generate_String_Wins(.Name, .PrizePoints, .Quotas, .Users, 1, TeamUser)
104         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(TextWinner, FontTypeNames.FONTTYPE_EVENT))
        
106         Call PrizeUser_All(SlotEvent, TeamUser)
108         Call CloseEvent(SlotEvent)
    
        End With
   
        '<EhFooter>
        Exit Sub

Events_Finish_Err:
        LogError Err.description & vbCrLf & _
               "in Events_Finish " & _
               "at line " & Erl

        '</EhFooter>
End Sub

' # Agrega un Round ganador al team seleccionado
Private Function Events_AddRound(ByVal SlotEvent As Byte, _
                            ByVal Team As Byte) As Integer
        '<EhHeader>
        On Error GoTo Events_AddRound_Err
        '</EhHeader>
                            
        Dim A As Long
        Dim Temp As Long
    
100     With Events(SlotEvent)
        
102         For A = LBound(.Users) To UBound(.Users)
104             If .Users(A).Team = Team Then
106                 .Users(A).RoundsWin = .Users(A).RoundsWin + 1
                    .Users(A).RoundsWinFinal = .Users(A).RoundsWinFinal + 1
108                 Temp = .Users(A).RoundsWin
                End If
110         Next A
        
112         Events_AddRound = Temp
        End With
        '<EhFooter>
        Exit Function

Events_AddRound_Err:
        LogError Err.description & vbCrLf & _
               "in Events_AddRound " & _
               "at line " & Erl

        '</EhFooter>
End Function

' # Busca a un oponente posible para seguir al evento

Private Function Event_SearchOponent(ByVal SlotEvent As Byte, ByVal Team As Byte) As Byte
        '<EhHeader>
        On Error GoTo Event_SearchOponent_Err
        '</EhHeader>
        Dim A As Long
    
100     With Events(SlotEvent)
102         For A = LBound(.Users) To UBound(.Users)
104             If .Modality = Enfrentamientos Then
106                 If .Users(A).ID > 0 And .Users(A).Oponent = Team Then
108                     Event_SearchOponent = .Users(A).Team
                        Exit Function
                    End If
                Else
110                 If .Users(A).ID > 0 And .Users(A).Team <> Team Then
112                     Event_SearchOponent = A
                        Exit Function
                    End If
                End If
114         Next A
        End With
    
        '<EhFooter>
        Exit Function

Event_SearchOponent_Err:
        LogError Err.description & vbCrLf & _
               "in Event_SearchOponent " & _
               "at line " & Erl

        '</EhFooter>
End Function



' [ESTRUCTURA]
' Ganadores
' ANCIENT (18pts) (+1.75%), LUCAS (22pts) (+2.12%) y ARNALDO (14pts)
' Ganador ANCIENT (18pts) (+1.75%)
' Aplica un porcentaje de 1 a 100 según el máximo recibido


Public Function Events_Generate_String_Wins(ByVal NameEvent As String, ByVal Points As Integer, _
                                            ByVal QuotasMax As Integer, ByRef Users() As tUserEvent, ByVal CountWin As Integer, ByVal TeamWin As Byte)

    On Error GoTo ErrHandler
    
    Dim A As Long
    Dim Temp As String
    Dim TempUser As String
    Dim Bonus As Single
    Dim Current As Integer
    Dim Lvl As Integer
    Dim PlayerRounds As Integer
    Dim FinalPoints As Integer
    Dim fragWeight As Single
    
    Temp = "«" & IIf(CountWin <> 1, "Ganadores", "Ganador") & " del evento " & NameEvent & "»" & vbCrLf
    Current = 0

    For A = LBound(Users) To UBound(Users)
        With Users(A)
            If .ID > 0 Then
                If .Team = TeamWin Then
                    Current = Current + 1
                    Lvl = UserList(.ID).Stats.Elv
                    
                    fragWeight = 1.4
                    Bonus = Porcentaje_Per_Level_Log(Lvl)
                    'PlayerRounds = .RoundsWinFinal
                    FinalPoints = Points * (1 + Bonus)
                    PrizeUser .ID, FinalPoints
                    TempUser = TempUser & Current & "°» " & UserList(.ID).Name & "  (Lvl" & Lvl & ") (" & Points & "pts»" & FinalPoints & "pts) +" & Transformar_En_Porcentaje(Bonus) & "% bonus" & vbCrLf
                End If
            End If
        End With
    Next A

    TempUser = Left$(TempUser, Len(TempUser) - 1)
    Temp = Temp & TempUser
    Events_Generate_String_Wins = Temp
    
    Exit Function
ErrHandler:
End Function



' # Eventos donde se catalogue como ganador al único existente.
Public Sub Events_CheckInscribed(ByVal UserIndex As Integer, _
                                 ByVal SlotEvent As Byte, _
                                 ByVal SlotEventUser As Byte, _
                                 ByVal TeamUser As Byte, _
                                 ByVal MapFight As Byte, _
                                 Optional ByVal Disconnect As Boolean = False)
        '<EhHeader>
        On Error GoTo Events_CheckInscribed_Err
        '</EhHeader>
    

    
        Dim SlotUser As Byte
        Dim Winner   As Integer
        Dim Oponent  As Integer
    
100     Oponent = Event_SearchOponent(SlotEvent, TeamUser)
    
102     With Events(SlotEvent)
104         Select Case .Modality
                Case eModalityEvent.DagaRusa, _
                   eModalityEvent.CastleMode, _
                   eModalityEvent.Busqueda, _
                   eModalityEvent.Unstoppable, _
                   eModalityEvent.JuegosDelHambre, _
                   eModalityEvent.DeathMatch
                
106                 If Oponent = 0 Then
108                     Call CloseEvent(SlotEvent)
                    Else
110                     If .Inscribed <= TeamCant(SlotEvent, Oponent) Then
112                         Call Events_Finish(SlotEvent, Oponent)
                        End If
                    End If
                
114             Case eModalityEvent.Enfrentamientos
                    Dim strWin    As String
                    Dim Text      As String
                    Dim UserCheck As Integer
                    Dim CantTeam As Integer
                    
116                 If Not Fight_CheckContinue(UserIndex, SlotEvent, TeamUser) Then
118                     UserCheck = Events_AddRound(SlotEvent, Oponent)
                            
120                     If UserCheck >= (.LimitRound / 2) + 0.5 Then
                        
CheckGanador:
122                         If MapFight > 0 Then

124                             MapEvent.Fight(MapFight).Run = False
126                             Call Fight_TeamWin(SlotEvent, Oponent, strWin)
128                             Text = Events(SlotEvent).Name & "» Gana " & strWin & "."
130                             Call Team_UserDie(SlotEvent, TeamUser, Events(SlotEvent).Users)
                                  
                                  If .Inscribed <= (.TeamCant * 2) Then
                                        .LimitRound = .LimitRoundFinal
                                  End If
                                  
                                  Events(SlotEvent).ArenasOcupadas = Events(SlotEvent).ArenasOcupadas - 1
                            End If
                            
132                         If .Inscribed <= TeamCant(SlotEvent, Oponent) Then
134                             ' Entrega el premio
                                  Text = Events_Generate_String_Wins(.Name, .PrizePoints, MaxKills(.TeamCant, .Quotas, .LimitRound, .LimitRoundFinal), .Users, .Inscribed, Oponent)
138                             Call CloseEvent(SlotEvent, Text)
                            Else
140                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Text, FontTypeNames.FONTTYPE_EVENT, eMessageType.cEvents_Curso))
142                             Call NewRound(SlotEvent)
144                             Call Fight_Combate(SlotEvent)
                            End If
                        Else

146                         If Not Fight_CheckContinue(UserIndex, SlotEvent, TeamUser, 1) Then
                               GoTo CheckGanador:
                            Else
148                             Fight_WarpTeam SlotEvent, MapFight, TeamUser, True, vbNullString, Oponent, False
150                             Fight_WarpTeam SlotEvent, MapFight, Oponent, False, vbNullString, TeamUser, False
                            End If
                        End If
                    
                    
                    End If
                
            End Select
        End With
    

        '<EhFooter>
        Exit Sub

Events_CheckInscribed_Err:
        LogError Err.description & vbCrLf & _
               "in Events_CheckInscribed " & _
               "at line " & Erl
        Call CloseEvent(SlotEvent)
        
        '</EhFooter>
End Sub

' # Eventos en los que el personaje se muere (proximamente tmb deslogeo todo en uno organizado)
Public Sub Events_UserDie(ByVal UserIndex As Integer, _
                            Optional ByVal AttackerIndex As Integer)
        '<EhHeader>
        On Error GoTo Events_UserDie_Err
        '</EhHeader>
        Dim SlotEvent As Byte
        Dim SlotEventUser As Byte
        Dim Temp As Integer
        Dim TeamUser As Integer
        Dim TeamOponent As Integer
        Dim MapFight As Byte
    
        Dim strWin As String
    
100     SlotEvent = UserList(UserIndex).flags.SlotEvent
102     SlotEventUser = UserList(UserIndex).flags.SlotUserEvent
104     TeamUser = Events(SlotEvent).Users(SlotEventUser).Team
106     TeamOponent = Events(SlotEvent).Users(SlotEventUser).Oponent
108     MapFight = Events(SlotEvent).Users(SlotEventUser).MapFight
    
110     Select Case Events(SlotEvent).Modality
            Case eModalityEvent.DeathMatch, eModalityEvent.GranBestia, eModalityEvent.JuegosDelHambre, eModalityEvent.DagaRusa
112             Call AbandonateEvent(UserIndex)
114             Call Events_CheckInscribed(UserIndex, SlotEvent, 0, TeamUser, 0)
        
116         Case eModalityEvent.Unstoppable
118             If AttackerIndex = 0 Then Exit Sub
120             If General.AntiFrags_CheckUser(AttackerIndex, UserIndex, 20) Then
122                 Events(SlotEvent).Users(UserList(AttackerIndex).flags.SlotUserEvent).Value = Events(SlotEvent).Users(UserList(AttackerIndex).flags.SlotUserEvent).Value + 1
                End If
            
124             Call Events_UserRevive(UserIndex)
            
126         Case eModalityEvent.CastleMode
128             Call Events_UserRevive(UserIndex)
        
130         Case eModalityEvent.Enfrentamientos
132             Call Events_CheckInscribed(UserIndex, SlotEvent, SlotEventUser, TeamUser, MapFight)
              
        End Select
        '<EhFooter>
        Exit Sub

Events_UserDie_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_UserDie " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Eventos en los que el personaje deba revivir y/o ser teletransportado
Private Sub Events_UserRevive(ByVal UserIndex As Integer)
    Dim Pos As WorldPos
    
    With UserList(UserIndex)
        .Counters.Shield = 5
        
        Call RevivirUsuario(UserIndex)
        
        Pos.Map = UserList(UserIndex).Pos.Map
        Pos.X = RandomNumber(.Pos.X - 7, .Pos.X + 7)
        Pos.Y = RandomNumber(.Pos.Y - 7, .Pos.Y + 7)
        
        Call EventWarpUser(UserIndex, Pos.Map, Pos.X, Pos.Y)
        
        
    End With
    
End Sub

' FIN EVENTO DAGA RUSA ###########################################
Private Function Events_ChangeBody() As Integer

    Dim Random As Integer
    Dim CharBody As Integer
    
    Random = RandomNumber(1, 8)
          
    Select Case Random

        Case 1 ' Zombie
                CharBody = 11

        Case 2 ' Golem
                CharBody = 11

        Case 3 ' Araña
                CharBody = 42

        Case 4 ' Asesino
                CharBody = 11 '48

        Case 5 'Medusa suprema
                CharBody = 151

        Case 6 'Dragón azul
                CharBody = 42 '247

        Case 7 'Viuda negra 185
                CharBody = 185

        Case 8 'Tigre salvaje
                CharBody = 147

    End Select

    Events_ChangeBody = CharBody
End Function


' # Se modifica la apariencia del personaje según elegido previamente.
Public Sub Events_ChangeApparience(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Events_ChangeApparience_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
102         Call ChangeBodyEvent(UserIndex, True)
104         .ShowName = False
        
        
        
106         If .flags.SlotEvent Then
108             If Events(.flags.SlotEvent).config(eConfigEvent.eMezclarApariencias) = 1 Then
110                 .Counters.TimeApparience = 1
                End If
            End If
        End With
        '<EhFooter>
        Exit Sub

Events_ChangeApparience_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_ChangeApparience " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Events_ChangePosition(ByVal UserIndex As Integer, ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo Events_ChangePosition_Err
        '</EhHeader>
    
        Dim Pos As WorldPos
    
        ' Para que cambiar con dos inscriptos
100     If Events(SlotEvent).Inscribed = TeamCant(SlotEvent, UserIndex) Then Exit Sub
    
102     Select Case Events(SlotEvent).Modality
            Case eModalityEvent.DeathMatch
104             Pos.Map = MapEvent.DeathMatch.Map
106             Pos.X = RandomNumber(MapEvent.DeathMatch.X - 10, MapEvent.DeathMatch.X + 10)
108             Pos.Y = RandomNumber(MapEvent.DeathMatch.Y - 10, MapEvent.DeathMatch.Y + 10)
            
110             If Events(SlotEvent).config(eConfigEvent.eTeletransportacion) = 1 Then
112                 UserList(UserIndex).Counters.TimeTelep = 15
                End If
        End Select
    
114     If Pos.Map Then Call EventWarpUser(UserIndex, Pos.Map, Pos.X, Pos.Y)
        '<EhFooter>
        Exit Sub

Events_ChangePosition_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_ChangePosition " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' DEATHMATCH ####################################################
Private Sub InitDeathMatch(ByVal SlotEvent As Byte)

    On Error GoTo error

    Dim LoopC As Integer

    Dim Pos   As WorldPos
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                .Users(LoopC).Team = LoopC
                .Users(LoopC).Selected = 1
                
                Call Events_ChangeApparience(.Users(LoopC).ID)
                Call Events_ChangePosition(.Users(LoopC).ID, SlotEvent)
            End If
              
        Next LoopC
          
        .TimeCount = 20
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : InitDeathMatch()"
End Sub

' EVENTO BUSQUEDA '
Private Sub InitBusqueda(ByVal SlotEvent As Byte)

    On Error GoTo error
          
    Dim LoopC As Integer

    Dim Pos   As WorldPos

    With Events(SlotEvent)
        
        ' Limpieza de objetos anteriores
        Call DeleteObjectMap(MapEvent.Busqueda.Map)
        
        ' Creación de objetos nuevos
        For LoopC = 1 To 20
            Call Create_ObjectMap(1, 1, MapEvent.Busqueda.Map, RandomNumber(20, 80), RandomNumber(20, 80), 1)
        Next LoopC
              
        ' Teletransportación de personajes a sus posiciones.
        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                Pos.Map = MapEvent.Busqueda.Map
                Pos.X = RandomNumber(50, 60)
                Pos.Y = RandomNumber(50, 60)
                      
                Call ClosestLegalPos(Pos, Pos)
                Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, True)
            End If

        Next LoopC
              
        .TimeFinish = 60
          
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : InitBusqueda()"
End Sub

Public Sub Busqueda_GetObj(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte)

    On Error GoTo error

    With Events(SlotEvent)
        .Users(SlotUserEvent).Value = .Users(SlotUserEvent).Value + 1
              
        WriteConsoleMsg .Users(SlotUserEvent).ID, "Has recolectado un objeto del piso. En total llevas " & .Users(SlotUserEvent).Value & " objetos recolectados. Sigue así!", FontTypeNames.FONTTYPE_INFO
        Create_ObjectMap 1, 1, MapEvent.Busqueda.Map, RandomNumber(30, 80), RandomNumber(30, 80), 1
    End With

    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Busqueda_GetObj()"
End Sub

' ENFRENTAMIENTOS ###############################################

Private Sub Fight_WarpTeam(ByVal SlotEvent As Byte, _
                           ByVal ArenaSlot As Byte, _
                           ByVal TeamEvent As Byte, _
                           ByVal IsContrincante As Boolean, _
                           ByRef StrTeam As String, _
                           ByVal TeamOponent As Integer, _
                           ByVal NewFight As Boolean)

    On Error GoTo error

    Dim LoopC   As Integer

    Dim strTemp As String, strTemp1 As String, strTemp2 As String

    Dim X       As Long, Y As Long
    
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 And .Users(LoopC).Team = TeamEvent Then
                LogEventos "Usuario: " & UserList(.Users(LoopC).ID).Name & ". Team: " & TeamEvent
                
                If .IsPlante > 0 Then
                    If IsContrincante Then
                        X = MapEvent.Fight(ArenaSlot).X + MapEvent.Fight(ArenaSlot).MAP_TILE_VS
                        Y = MapEvent.Fight(ArenaSlot).Y
                    Else
                        X = MapEvent.Fight(ArenaSlot).X
                        Y = MapEvent.Fight(ArenaSlot).Y

                    End If
                    
                Else

                    If IsContrincante Then
                        X = MapEvent.Fight(ArenaSlot).X + MapEvent.Fight(ArenaSlot).MAP_TILE_VS
                        Y = MapEvent.Fight(ArenaSlot).Y + MapEvent.Fight(ArenaSlot).MAP_TILE_VS
                    Else
                        X = MapEvent.Fight(ArenaSlot).X
                        Y = MapEvent.Fight(ArenaSlot).Y

                    End If

                End If
                
                Call EventWarpUser(.Users(LoopC).ID, MapEvent.Fight(ArenaSlot).Map, X, Y)
                
                If IsContrincante Then
                    UserList(.Users(LoopC).ID).flags.FightTeam = 2
                    RefreshCharStatus (.Users(LoopC).ID)
                Else
                    UserList(.Users(LoopC).ID).flags.FightTeam = 1
                    RefreshCharStatus (.Users(LoopC).ID)
                End If
                      
                If StrTeam = vbNullString Then
                    StrTeam = UserList(.Users(LoopC).ID).Name
                Else
                    StrTeam = StrTeam & "-" & UserList(.Users(LoopC).ID).Name

                End If
                
                .Users(LoopC).Oponent = TeamOponent
                .Users(LoopC).Value = 1
                .Users(LoopC).MapFight = ArenaSlot
                      
                UserList(.Users(LoopC).ID).Counters.TimeFight = 10
                Call WriteUserInEvent(.Users(LoopC).ID)
                Call Events_StatsFull(.Users(LoopC).ID, NewFight)

            End If

        Next LoopC
        
        
        
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Fight_WarpTeam()"
    Call CloseEvent(SlotEvent)

End Sub

Private Function Fight_Search_Enfrentamiento(ByVal UserIndex As Integer, _
                                             ByVal UserTeam As Byte, _
                                             ByVal SlotEvent As Byte) As Byte

    On Error GoTo error

    ' Chequeamos que tengamos contrincante para luchar.
    Dim LoopC As Integer
          
    Fight_Search_Enfrentamiento = 0
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 And .Users(LoopC).Value = 0 Then
                If .Users(LoopC).ID <> UserIndex And .Users(LoopC).Team <> UserTeam Then
                    Fight_Search_Enfrentamiento = .Users(LoopC).Team

                    Exit For

                End If
            End If

        Next LoopC
          
    End With
          
    Exit Function

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Fight_Search_Enfrentamiento()"
    Call CloseEvent(SlotEvent)
End Function

Private Sub NewRound(ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo NewRound_Err
        '</EhHeader>

        Dim LoopC As Long

        Dim Count As Long

100     With Events(SlotEvent)
102         Count = 0
              
104         For LoopC = LBound(.Users()) To UBound(.Users())

106             If .Users(LoopC).ID > 0 Then

                    ' Hay esperando
108                 If .Users(LoopC).Value = 0 Then

                        Exit Sub

                    End If
                      
                    ' Hay luchando
110                 If .Users(LoopC).MapFight > 0 Then

                        Exit Sub

                    End If
                End If

112         Next LoopC
              
114         For LoopC = LBound(.Users()) To UBound(.Users())
116             .Users(LoopC).Value = 0
118             .Users(LoopC).RoundsWin = 0
                .Users(LoopC).Damage = 0
                .Users(LoopC).TimeCancel = 60
120         Next LoopC


              
122         LogEventos "Se reinicio la informacion de los fights()"
              
        End With
    
    

        '<EhFooter>
        Exit Sub

NewRound_Err:
        LogError Err.description & vbCrLf & _
               "in NewRound " & _
               "at line " & Erl
        
        Call CloseEvent(SlotEvent)
        '</EhFooter>
End Sub

Private Sub Fight_Combate(ByVal SlotEvent As Byte)

    On Error GoTo error

    ' Buscamos una arena disponible y mandamos la mayor cantidad de usuarios disponibles.
    Dim LoopC       As Integer

    Dim FreeArena   As Byte

    Dim OponentTeam As Byte

    Dim strTemp     As String

    Dim strTeam1    As String

    Dim strTeam2    As String
          
    With Events(SlotEvent)
cheking:

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 And .Users(LoopC).Value = 0 Then
                FreeArena = FreeSlotArena(SlotEvent)
                      
                If FreeArena > 0 Then
                    OponentTeam = Fight_Search_Enfrentamiento(.Users(LoopC).ID, .Users(LoopC).Team, SlotEvent)
                          
                    If OponentTeam > 0 Then
                        
                        .Users(LoopC).TimeCancel = 60
                        
                        Fight_WarpTeam SlotEvent, FreeArena, .Users(LoopC).Team, False, strTeam1, OponentTeam, True
                        Fight_WarpTeam SlotEvent, FreeArena, OponentTeam, True, strTeam2, .Users(LoopC).Team, True
                        MapEvent.Fight(FreeArena).Run = True
                            
                        
                        strTemp = "Torneo " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» "
                        strTemp = strTemp & strTeam1 & " vs " & strTeam2
                        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_EVENT, eMessageType.cEvents_Curso)
                              
                        strTemp = vbNullString
                        strTeam1 = vbNullString
                        strTeam2 = vbNullString
                              
                    Else
                        ' Pasa de ronda automaticamente
                        .Users(LoopC).Value = 1
                        WriteConsoleMsg .Users(LoopC).ID, "Hemos notado que no tienes un adversario. Pasaste a la siguiente ronda.", FontTypeNames.FONTTYPE_INFO
                        NewRound SlotEvent
                        GoTo cheking:
                    End If
                End If
            End If

        Next LoopC
              
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Fight_Combate()"
    Call CloseEvent(SlotEvent)
End Sub

Private Function CheckTeam_UserDie(ByVal SlotEvent As Integer, _
                                   ByVal TeamUser As Byte) As Boolean

    On Error GoTo error

    Dim LoopC As Integer

    ' Encontramos a uno del Team vivo, significa que no hay terminación del duelo.
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                If .Users(LoopC).Team = TeamUser Then
                    If UserList(.Users(LoopC).ID).flags.Muerto = 0 Then
                        CheckTeam_UserDie = False

                        Exit Function

                    End If
                End If
            End If

        Next LoopC
              
        CheckTeam_UserDie = True
          
    End With
          
    Exit Function

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : CheckTeam_UserDie()"
End Function
Private Sub Team_UserDie(ByVal SlotEvent As Byte, _
                         ByVal TeamSlot As Byte, _
                         ByRef Users() As tUserEvent)
        '<EhHeader>
        On Error GoTo Team_UserDie_Err
        '</EhHeader>

    
        Dim LoopC As Integer
        Dim UserIndex As Integer
    
    
100     For LoopC = LBound(Users()) To UBound(Users())
102         If Not Events(SlotEvent).Run Then Exit Sub ' Si ya cerró el evento significa que no hace falta seguir recorriendo.
        
104         If Users(LoopC).ID > 0 Then
106             If Users(LoopC).Team = TeamSlot Then
108                 UserIndex = Users(LoopC).ID
110                 AbandonateEvent UserIndex
                End If
            End If

112     Next LoopC

        '<EhFooter>
        Exit Sub

Team_UserDie_Err:
        LogError Err.description & vbCrLf & _
               "in Team_UserDie " & _
               "at line " & Erl

        '</EhFooter>
End Sub
Public Function Fight_CheckContinue(ByVal UserIndex As Integer, _
                                    ByVal SlotEvent As Byte, _
                                    ByVal TeamSlot As Byte, _
                                    Optional ByVal Muerto As Byte = 0) As Boolean
        ' Esta función devuelve un TRUE cuando el enfrentamiento puede seguir.
        '<EhHeader>
        On Error GoTo Fight_CheckContinue_Err
        '</EhHeader>
          
        Dim LoopC As Integer, cant As Integer
          
100     With Events(SlotEvent)
              
102         Fight_CheckContinue = False
              
104         For LoopC = LBound(.Users()) To UBound(.Users())


            
                ' User válido
106             If .Users(LoopC).ID > 0 And (.Users(LoopC).ID <> UserIndex Or Muerto = 1) Then
108                 If .Users(LoopC).Team = TeamSlot Then
110                     If (UserList(.Users(LoopC).ID).flags.Muerto = 0) Or Muerto = 1 Then
112                         Fight_CheckContinue = True

                            Exit For

                        End If
                    End If
                End If

114         Next LoopC

        End With

        '<EhFooter>
        Exit Function

Fight_CheckContinue_Err:
        LogError Err.description & vbCrLf & _
               "in Fight_CheckContinue " & _
               "at line " & Erl
End Function

Private Sub Events_StatsFull(ByVal UserIndex As Integer, ByVal NewFight As Boolean)

    On Error GoTo error

    With UserList(UserIndex)

        If .flags.Muerto Then Call RevivirUsuario(UserIndex)
              
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinAGU = 100
        .Stats.MinHam = 100
        
        WriteUpdateUserStats UserIndex
          
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Events_StatsFull()"
End Sub


Public Sub Fight_TeamWin(ByVal SlotEvent As Byte, _
                         ByVal TeamSlotOponent As Byte, _
                         ByRef strTempWin As String)
        '<EhHeader>
        On Error GoTo Fight_TeamWin_Err
        '</EhHeader>


        Dim LoopC      As Integer

        Dim MapFight   As Byte
    
        Dim TeamWin As Byte
        
        Dim RoundsWin As Byte
        
100     With Events(SlotEvent)

        
102         For LoopC = LBound(.Users()) To UBound(.Users())
            
104             If .Users(LoopC).ID > 0 Then

106                 With .Users(LoopC)

108                     If .Team = TeamSlotOponent Then
110                         Events_StatsFull .ID, False

112                         If strTempWin = vbNullString Then
114                             strTempWin = UserList(.ID).Name
                            Else
116                             strTempWin = strTempWin & "-" & UserList(.ID).Name
                            End If
                            
                            RoundsWin = .RoundsWin
118                         MapFight = .MapFight
                             
120                         .Oponent = 0
122                         .MapFight = 0
                        
124                         EventWarpUser .ID, MapEvent.SalaEspera.Map, MapEvent.SalaEspera.X, MapEvent.SalaEspera.Y
126                         WriteConsoleMsg .ID, "Felicitaciones. Has ganado el duelo", FontTypeNames.FONTTYPE_INFO
                              
128                         UserList(.ID).flags.FightTeam = 0
130                         RefreshCharStatus (.ID)
                              
132                         If UserList(.ID).flags.Muerto Then RevivirUsuario (.ID)
                        End If

                    End With

                End If

134         Next LoopC
        
136
             strTempWin = strTempWin & " (Rounds: " & RoundsWin & ")"
        
        End With
        '<EhFooter>
        Exit Sub

Fight_TeamWin_Err:
        LogError Err.description & vbCrLf & _
               "in Fight_TeamWin " & _
               "at line " & Erl

        '</EhFooter>
End Sub

Private Function TeamCant(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte) As Byte

    On Error GoTo error

    ' Devuelve la cantidad de miembros que tiene un clan
    Dim LoopC As Integer
          
    TeamCant = 0
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).ID > 0 Then
                If .Users(LoopC).Team = TeamSlot Then
                    TeamCant = TeamCant + 1
                End If
            End If
            
        Next LoopC

    End With
          
    Exit Function

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : TeamCant()"
End Function

Private Function Events_StrReward(ByVal SlotEvent As Byte) As String
        '<EhHeader>
        On Error GoTo Events_StrReward_Err
        '</EhHeader>

        Dim Temp As String
    
100     With Events(SlotEvent)
102         If .PrizeGld = 0 And .PrizeEldhir = 0 And .PrizePoints = 0 And .PrizeObj.ObjIndex = 0 Then
104             Temp = "Premio entregado: NINGUNO"
                Exit Function
            Else
106             Temp = "Premio entregado: "
            
108             If .PrizeGld > 0 Then
110                 Temp = Temp & .PrizeGld & " Monedas de Oro. "
                End If
                    
112             If .PrizeEldhir > 0 Then
114                 Temp = Temp & .PrizeEldhir & " DSP. "
                End If
                
                If .PrizePoints > 0 Then
                 Temp = Temp & .PrizePoints & " Puntos de Torneo. "
                End If
                    
116             If .PrizeObj.ObjIndex > 0 Then
118                 Temp = Temp & ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ")"
                End If
            End If
        
        
120         Events_StrReward = Temp
        End With
    
        '<EhFooter>
        Exit Function

Events_StrReward_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_StrReward " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


' ############################## USUARIO UNSTOPPABLE ###########################################
Public Sub InitUnstoppable(ByVal SlotEvent As Byte)

    On Error GoTo error

    Dim LoopC As Integer
          
    With Events(SlotEvent)

        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).ID > 0 Then
                EventWarpUser .Users(LoopC).ID, MapEvent.Imparable.Map, RandomNumber(MapEvent.Imparable.X - 5, MapEvent.Imparable.X + 5), RandomNumber(MapEvent.Imparable.Y - 5, MapEvent.Imparable.Y + 5)
                      
            End If

        Next LoopC
              
        .TimeCount = 10
        .TimeFinish = 420 + .TimeCount
    End With
          
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : InitUnstoppable()"
End Sub


Public Function NewEvent_Configuration(ByVal Name As String, _
                                       ByVal Modality As eModalityEvent, _
                                       ByVal QuotasMin As Byte, _
                                       ByVal QuotasMax As Byte, _
                                       ByVal LvlMin As Byte, _
                                       ByVal LvlMax As Byte, _
                                       ByVal TimeInit As Integer, _
                                       ByVal TimeCancel As Integer, _
                                       ByVal PrizeGld As Long, _
                                       ByVal PrizeGldPremium As Long, _
                                       ByVal ObjIndex As Integer, _
                                       ByVal ObjAmount As Integer, _
                                       ByVal LimitRed As Integer, _
                                       ByRef config() As Byte, _
                                       ByRef AllowedClass() As Byte, _
                                       ByRef AllowedFaction() As Byte) As tEvents
        '<EhHeader>
        On Error GoTo NewEvent_Configuration_Err
        '</EhHeader>
                                       
                                                                       
        Dim Temp As tEvents
        Dim A As Long
    
100     With Temp
102         .QuotasMin = QuotasMin
              .QuotasMax = QuotasMax
104         .TimeInit = TimeInit
106         .TimeCancel = TimeCancel
108         .Modality = Modality
110         .Name = Name
        
112         .PrizeGld = PrizeGld
114         .PrizeEldhir = PrizeGldPremium
116         .PrizeObj.ObjIndex = ObjIndex
118         .PrizeObj.Amount = ObjAmount
        
122         .LvlMin = LvlMin
124         .LvlMax = LvlMax
        
126         .AllowedClasses = AllowedClass
128         .AllowedFaction = AllowedFaction
        
130         For A = 1 To MAX_EVENTS_CONFIG
132             .config(A) = config(A)
134         Next A
        
        End With
                                    
136     NewEvent_Configuration = Temp
        '<EhFooter>
        Exit Function

NewEvent_Configuration_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.NewEvent_Configuration " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


' Devolvemos la lista de participantes en un Texto separado por comas
Public Function Event_Text_Users(ByVal Slot As Byte) As String

    On Error GoTo ErrHandler

    Dim Modality As eModalityEvent

    Dim A        As Long

    Dim Temp     As String
    
    With Events(Slot)

        For A = 1 To .Quotas

            If .Users(A).ID > 0 Then
                If A = .Quotas Then
                    
                    Temp = Temp & " y " & UserList(.Users(A).ID).Name & "."
                Else
                    Temp = Temp & UserList(.Users(A).ID).Name & ", "
                    
                    If A + 1 = .Quotas Then
                        Temp = Left$(Temp, Len(Temp) - 2)
                    End If
                End If
            End If

        Next A
                
        
        Event_Text_Users = Temp
        
    End With
    
    Exit Function

ErrHandler:
End Function

Public Function Event_Text_Users_VS(ByVal Slot As Byte) As String
        '<EhHeader>
        On Error GoTo Event_Text_Users_VS_Err
        '</EhHeader>

        Dim A     As Long

        Dim TeamA As String

        Dim TeamB As String
    
100     With Events(Slot)

102         Select Case .Modality

                Case eModalityEvent.CastleMode

104                 For A = 1 To .Quotas

106                     If .Users(A).ID > 0 Then
108                         If A > (.Quotas / 2) Then
110                             TeamA = TeamA & UserList(.Users(A).ID).Name & ", "
                            Else
112                             TeamB = TeamB & UserList(.Users(A).ID).Name & ", "
                            End If
                        End If

114                 Next A
                
116                 TeamA = Left$(TeamA, Len(TeamA) - 2)
118                 TeamB = Left$(TeamB, Len(TeamB) - 2)
120                 Event_Text_Users_VS = TeamA & " VS " & TeamB
            
122             Case eModalityEvent.DeathMatch

124                 For A = 1 To .Quotas
126                     TeamA = TeamA & UserList(.Users(A).ID).Name & " VS "
128                 Next A
                
130                 TeamA = Left$(TeamA, Len(TeamA) - 4)
132                 Event_Text_Users_VS = TeamA
            End Select

        End With
    
        '<EhFooter>
        Exit Function

Event_Text_Users_VS_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_Text_Users_VS " & _
               "at line " & Erl
        
        '</EhFooter>
End Function



'#################################
' EVENTO DE TELEPORTS
'#################################

Public Sub Events_Teleports_Init(ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo Events_Teleports_Init_Err
        '</EhHeader>
    
        Dim X             As Long, Y As Long

        Dim A             As Long, ElectionX As Long, Election As Boolean
    
100
        
        Dim MapIndex As Integer
        MapIndex = Events_Teleports_SearchMapFree(Events(SlotEvent).Inscribed)
        
        If MapIndex = 0 Then
            Call CloseEvent(SlotEvent, , True)
            Exit Sub
        End If
        
        With MapEvent.Teleport(MapIndex)
            .Usage = True
            Y = .YInitial_TP
            
            ' Seteo todos los portales a la posición de comienzo
            For A = 1 To .Y_Pasajes
                For X = .XInitial_TP To (.XInitial_TP + .XTiles_TP)
                    With MapData(.Map, X, Y)
                
                        If .Blocked = 0 Then
                            .TileExit.Map = MapEvent.Teleport(MapIndex).Map
                            .TileExit.X = MapEvent.Teleport(MapIndex).XWarp
                            .TileExit.Y = MapEvent.Teleport(MapIndex).YWarp
                        End If
                
                    End With
                Next X
                
                Y = Y - .Y_TileAdd
            Next A
            
            Y = .YInitial_TP
            
            ' Determino el Portal Correcto
            For A = 1 To .Y_Pasajes
                Do While Election = False
                    ElectionX = RandomNumber(.XInitial_TP, .XInitial_TP + .XTiles_TP)
                    
                    With MapData(.Map, ElectionX, Y)
                        
                        If .Blocked = 0 Then
                             .TileExit.Map = MapEvent.Teleport(MapIndex).Map
                             .TileExit.X = MapEvent.Teleport(MapIndex).XWarp
                             .TileExit.Y = Y - 2
                             Election = True
                        End If
        
                    End With
                   
                Loop
        
                Y = Y - .Y_TileAdd
                Election = False
            Next A
            
        End With
        
148     With Events(SlotEvent)

150         For A = LBound(.Users()) To UBound(.Users())

152             If .Users(A).ID > 0 Then
154                 EventWarpUser .Users(A).ID, MapEvent.Teleport(MapIndex).Map, MapEvent.Teleport(MapIndex).XWarp, MapEvent.Teleport(MapIndex).YWarp
                End If

156         Next A

        End With
    
        '<EhFooter>
        Exit Sub

Events_Teleports_Init_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_Teleports_Init " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Events_Teleports_Finish(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Events_Teleports_Finish_Err
        '</EhHeader>

    
100     With UserList(UserIndex)

102         Call WriteConsoleMsg(UserIndex, "¡¡Felicitaciones!! Has ganado el Evento de Teleports.", FontTypeNames.FONTTYPE_CONSEJOCAOS)
104         Call PrizeUser(UserIndex, 0)
106         Call CloseEvent(.flags.SlotEvent, strModality(.flags.SlotEvent, eModalityEvent.Teleports) & "» El personaje " & .Name & " ha encontrado el Teleport Final. ¡¡Felicitaciones!!. " & Events_StrReward(.flags.SlotEvent))
    
        End With

        '<EhFooter>
        Exit Sub

Events_Teleports_Finish_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_Teleports_Finish " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



Public Sub Events_GranBestia_MuereNpc(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Events_GranBestia_MuereNpc_Err
        '</EhHeader>
    
        Dim SlotEvent  As Byte

        Dim strTemp(1) As String
    
100     SlotEvent = UserList(UserIndex).flags.SlotEvent
    
102     With Events(SlotEvent)
104         Call PrizeUser(UserIndex, 0)
106         strTemp(0) = strModality(SlotEvent, eModalityEvent.GranBestia) & "» El personaje " & UserList(UserIndex).Name & " ha acabado con la Gran Bestia. " & Events_StrReward(SlotEvent)
        
108         Call CloseEvent(SlotEvent, strTemp(0))
        End With

        '<EhFooter>
        Exit Sub

Events_GranBestia_MuereNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_GranBestia_MuereNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Events_GranBestia_MuereUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Events_GranBestia_MuereUser_Err
        '</EhHeader>

        Dim SlotEvent  As Byte

        Dim UserWinner As Integer

        Dim strTemp(1) As String
    
100     SlotEvent = UserList(UserIndex).flags.SlotEvent
    
102     With Events(SlotEvent)
104         Call AbandonateEvent(UserIndex)
        
            'If .Inscribed = 1 Then
              '  UserWinner = SearchLastUserEvent(SlotEvent)
            
                'Call PrizeUser(UserWinner, strTemp(1))
               ' strTemp(0) = strModality(SlotEvent, eModalityEvent.GranBestia) & "» El personaje " & UserList(UserWinner).Name & " ha logrado sobrevivir a la Gran Bestia. " & strTemp(1)
        
               ' Call CloseEvent(SlotEvent, strTemp(0))
           ' End If
    
        End With

        '<EhFooter>
        Exit Sub

Events_GranBestia_MuereUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_GranBestia_MuereUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



'#################################
' JUEGOS DEL HAMBRE (JDH)
'#################################

Private Sub Events_JDH_Init(ByVal SlotEvent As Byte)
    
    On Error GoTo ErrHandler
    
    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    Dim A As Long, B As Long
    Dim X As Long
    Dim Y As Long
    
    
    Pos.Map = MapEvent.JuegosDelHambre.Map
    
    Call Remove_All_Map(Pos.Map, 1, 1)
        
    For A = 1 To Int((Events(SlotEvent).Quotas) / 2)
        For B = 0 To 3
            Pos.X = RandomNumber(25, 75) ' Rectangulo de Combate
            Pos.Y = RandomNumber(15, 70) ' Rectangulo de Combate
            
            Call SpawnNpc(29 + B, Pos, False, False)
            
        Next B
        
        For B = 0 To 7
            Call Create_ObjectMap(629, 3, Pos.Map, RandomNumber(25, 75), RandomNumber(15, 70), 1)
        Next B
    Next A
    
    X = 41
    Y = 82
    
    For A = LBound(Events(SlotEvent).Users()) To UBound(Events(SlotEvent).Users())

        If Events(SlotEvent).Users(A).ID > 0 Then
            Pos.X = X
            Pos.Y = Y
            
            Call WarpUserChar(Events(SlotEvent).Users(A).ID, Pos.Map, Pos.X, Pos.Y, True)
            
            X = X + 1
        End If

        Next A
    
    Exit Sub
ErrHandler:

End Sub








' # Devuelve el UserIndex que quedó como único inscripto.
Private Function SearchLastUserEvent(ByVal SlotEvent As Byte) As Integer
        '<EhHeader>
        On Error GoTo SearchLastUserEvent_Err
        '</EhHeader>

        ' Busca el último usuario que está en el torneo. En todos los eventos será el ganador.
          
        Dim LoopC As Integer
          
100     With Events(SlotEvent)

102         For LoopC = LBound(.Users()) To UBound(.Users())

104             If .Users(LoopC).ID > 0 Then
106                 SearchLastUserEvent = .Users(LoopC).ID

                    Exit Function

                End If

108         Next LoopC

        End With

        '<EhFooter>
        Exit Function

SearchLastUserEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.SearchLastUserEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Events_OrdenateUsersValue(ByVal SlotEvent As Byte, _
                                     ByRef CopyUsers() As tUserEvent)
On Error GoTo error
    Dim A    As Long, B As Long
    Dim Temp As tUserEvent
    
    With Events(SlotEvent)
    
        ReDim CopyUsers(LBound(.Users()) To UBound(.Users())) As tUserEvent
            
        For B = LBound(.Users()) To UBound(.Users())
            CopyUsers(B) = .Users(B)
        Next B
        
    
        For A = LBound(.Users()) To UBound(.Users())
            For B = LBound(.Users()) To UBound(.Users()) - A
                If CopyUsers(B).Value < CopyUsers(B + 1).Value Then
                    Temp = CopyUsers(B)
                    CopyUsers(B) = CopyUsers(B + 1)
                    CopyUsers(B + 1) = Temp
                End If
            Next B
        Next A
    End With
    
    Exit Sub

error:
    LogEventos "[" & Err.number & "] " & Err.description & ") PROCEDIMIENTO : Events_OrdenateUsersValue()"
                
End Sub

' # Generamos la tabla de posiciones.
Private Function Event_GenerateTablaPos(ByVal SlotEvent As Byte, _
                                        ByRef CopyUsers() As tUserEvent) As String
        '<EhHeader>
        On Error GoTo Event_GenerateTablaPos_Err
        '</EhHeader>

        Dim LoopC As Integer
          
100     With Events(SlotEvent)

102         For LoopC = LBound(.Users()) To UBound(.Users())

104             If CopyUsers(LoopC).ID > 0 Then
106                 Event_GenerateTablaPos = Event_GenerateTablaPos & LoopC & "° »» " & UserList(CopyUsers(LoopC).ID).Name & " (" & CopyUsers(LoopC).Value & ")" & vbCrLf
                End If

108         Next LoopC

        End With
          
        '<EhFooter>
        Exit Function

Event_GenerateTablaPos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_GenerateTablaPos " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' # Se mezclan los personajes del evento.
Public Sub Event_RandomUsers_Array(ByVal Slot As Byte, ByRef vArray() As tUserEvent)
        '<EhHeader>
        On Error GoTo Event_RandomUsers_Array_Err
        '</EhHeader>
      
        Dim i          As Long

        Dim rndIndex   As Long

        Dim Temp       As tUserEvent

        Dim startIndex As Integer

        Dim endIndex   As Integer
    
      
100     startIndex = LBound(vArray)
102     endIndex = UBound(vArray)
      
104     For i = startIndex To endIndex
106         rndIndex = Int((endIndex - startIndex + 1) * Rnd() + startIndex)
  
108         Temp = vArray(i)
110         vArray(i) = vArray(rndIndex)
112         vArray(rndIndex) = Temp
        
114         With Events(Slot)

116             If .Users(rndIndex).ID > 0 Then
118                 UserList(.Users(rndIndex).ID).flags.SlotUserEvent = rndIndex
                End If
            
120             If .Users(i).ID > 0 Then
122                 UserList(.Users(i).ID).flags.SlotUserEvent = i
                End If

            End With

124     Next i
    
        '<EhFooter>
        Exit Sub

Event_RandomUsers_Array_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Event_RandomUsers_Array " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Se reinician TODOS los eventos en curso.
Public Sub Eventos_Reset_All()
        '<EhHeader>
        On Error GoTo Eventos_Reset_All_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_EVENT_SIMULTANEO

102         With Events(A)

104             If .Run Then
106                 Call CloseEvent(A, , True)
                End If

            End With

108     Next A

        '<EhFooter>
        Exit Sub

Eventos_Reset_All_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Eventos_Reset_All " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Private Sub AssignTeam(ByVal SlotEvent As Byte)
    On Error GoTo AssignTeam_Err

    Dim A As Integer
    Dim UserIndex As Integer
    Dim Team As Byte
    Team = 1

    Dim ProcessedGroups As Collection
    Set ProcessedGroups = New Collection
    Dim IndividualCounter As Integer
    IndividualCounter = 0
10
    With Events(SlotEvent)
        For A = LBound(.Users) To UBound(.Users)
            If .Users(A).Name <> vbNullString Then
                UserIndex = NameIndex(.Users(A).Name)
20
                If UserIndex > 0 Then
                
                    If .TeamCant <> 1 Then
                        If UserList(UserIndex).GroupIndex > 0 Then
                            ' Si el usuario tiene un grupo y no ha sido procesado aún
                            If Not GroupProcessed(UserList(UserIndex).GroupIndex, ProcessedGroups) Then
                                ' Marcar el grupo como procesado
                                ProcessedGroups.Add UserList(UserIndex).GroupIndex, CStr(UserList(UserIndex).GroupIndex)
25
                                ' Asignar todos los miembros de este grupo al mismo equipo
                                AssignGroupToTeam UserList(UserIndex).GroupIndex, Team, Events(SlotEvent)
30
                                ' Incrementar el número de equipo
                                Team = Team + 1
                            End If
                        Else
                            ' Si el usuario no tiene grupo, asignarle un equipo en función del módulo de la cantidad total de equipos
                            .Users(A).Team = (Team - 1) Mod .TeamCant + 1
                            
                            ' Incrementar el contador individual
                            IndividualCounter = IndividualCounter + 1
                            
                            ' Si se ha asignado un equipo a la cantidad deseada de usuarios individuales, aumenta el número de equipo
                            If IndividualCounter Mod .TeamCant = 0 Then
                                Team = Team + 1
                            End If
                            
                            
                        End If
                    Else
                        .Users(A).Team = Team
                        Team = Team + 1
                    End If
                    
                End If
            End If
            
        Next A
    End With
75
    Exit Sub

AssignTeam_Err:
    LogError Err.description & vbCrLf & _
           "in AssignTeam " & _
           "at line " & Erl
    Call CloseEvent(SlotEvent)
End Sub

Private Function GroupProcessed(ByVal GroupIndex As Integer, ByVal ProcessedGroups As Collection) As Boolean
    On Error Resume Next
    Dim Temp As Variant
    Temp = ProcessedGroups.Item(CStr(GroupIndex))
    If Err.number = 0 Then
        GroupProcessed = True
    Else
        GroupProcessed = False
    End If
    On Error GoTo 0
End Function

Private Sub AssignGroupToTeam(ByVal GroupIndex As Integer, ByVal Team As Byte, ByRef EventData As tEvents)

    On Error GoTo ErrHandler
    
    Dim A As Integer
    Dim tUser As Integer
    
    With Groups(GroupIndex)
        For A = LBound(.User) To UBound(.User)
            tUser = .User(A).Index
            EventData.Users(UserList(tUser).flags.SlotUserEvent).Team = Team
        Next A
        
    End With
    
ErrHandler:
    Exit Sub
     LogError Err.description & vbCrLf & "in AssignGroupToTeam " & "at line " & Erl
    
End Sub
Private Sub RemoveGroupToTeam(ByVal GroupIndex As Integer, ByRef EventData As tEvents)

    On Error GoTo ErrHandler
    Dim A As Integer
    Dim tUser As Integer
    
    With Groups(GroupIndex)
        For A = LBound(.User) To UBound(.User)
            tUser = .User(A).Index
            EventData.Users(UserList(tUser).flags.SlotUserEvent).Name = vbNullString
        Next A
        
    End With
    
ErrHandler:
    Exit Sub
     LogError Err.description & vbCrLf & "in RemoveGroupToTeam " & "at line " & Erl
End Sub

' # Comprueba el evento y en base a los cupos lo completa.
Public Function Events_UpdateQuotas(ByVal SlotEvent As Byte) As Boolean

        '<EhHeader>
        On Error GoTo Events_UpdateQuotas_Err

        '</EhHeader>
    
100     With Events(SlotEvent)
            
            
            Dim Intentos As Integer
            Dim Inscriptos As Integer
            Dim tUser() As Integer
110
            ' # Asigna el Team a los participantes
            Call AssignTeam(SlotEvent)
120
            ' # No está siendo llenado, esperamos otro tiempo
            If .Inscribed < .QuotasMin Then
                Exit Function
            End If
130
            Do While (.Inscribed > Inscriptos + (.QuotasMin * 2))
                Inscriptos = Inscriptos + (.QuotasMin * 2)
                Intentos = Intentos + 1
                
                 If Intentos >= 30 Then Exit Function
            Loop
140
            If Inscriptos = 0 Then Inscriptos = .Inscribed
            ReDim tUser(1 To Inscriptos) As Integer
            ' Actualiza lista de disponibles
            Call Events_CheckUsers(SlotEvent, tUser, Inscriptos)
150
             ' # No está siendo llenado, esperamos otro tiempo
            If .Inscribed < .QuotasMin Then
                Exit Function
            End If
160
            ' Crea la participación de los selectos.
            Call Events_UpdateUsers(SlotEvent, Inscriptos, tUser)
170
          End With
          
180     Events_UpdateQuotas = True
        '<EhFooter>
        Exit Function

Events_UpdateQuotas_Err:
        LogError Err.description & vbCrLf & "in Events_UpdateQuotas " & "at line " & Erl

        '</EhFooter>
End Function
'# Comprueba si los personajes anotados pueden jugar, si no los quitamos de la lista antes de comprobar si se puede disfrutar del evento o se esperan 3 minutos más.
Public Sub Events_CheckUsers(ByVal SlotEvent As Byte, ByRef Users() As Integer, ByVal Inscriptos As Integer)
    Dim A As Long
    Dim tUser As Integer
    Dim ErrorMsg As String
    
    Dim TeamMemberFailed As Boolean ' Booleano para verificar si algún miembro del equipo falla
    Dim Slot As Byte
    
    With Events(SlotEvent)
    

        For A = LBound(.Users) To UBound(.Users)
            If .Users(A).Name <> vbNullString Then
                tUser = NameIndex(.Users(A).Name)
                
                ' # Comprueba los usuarios que deslogearon luego de la inscripción y los retiramos de la lista.
                If tUser = 0 Then
                    .Users(A).Name = vbNullString
                    .Inscribed = .Inscribed - 1
                Else
                    ' # Comprueba que cumpla los requisitos
                    If Not Events_CheckUserEvent(tUser, SlotEvent, ErrorMsg) Then
                        ' Si un miembro del equipo no puede participar, todo el equipo queda fuera
                        If UserList(tUser).GroupIndex > 0 Then
                            Call RemoveGroupToTeam(UserList(tUser).GroupIndex, Events(SlotEvent))
                        Else
                            .Users(A).Name = vbNullString
                            .Inscribed = .Inscribed - 1
                        End If
                    Else
                        If .Inscribed > Inscriptos Then
                            ' Si un miembro del equipo no puede participar, todo el equipo queda fuera
                            If UserList(tUser).GroupIndex > 0 Then
                                Call RemoveGroupToTeam(UserList(tUser).GroupIndex, Events(SlotEvent))
                            Else
                                .Users(A).Name = vbNullString
                                .Inscribed = .Inscribed - 1
                            End If
                        Else
                            ' # Lo guarda en el array de usuarios conectados
                            Slot = Slot + 1
                            Users(Slot) = tUser
                        End If
                    End If
                End If
            End If
        Next A
    End With
End Sub


' # Resetea la PRE-Inscripción

' # Actualizamos y expulsamos a los que no tengan que estar por cantidad de cupos reducida/ajustada.
Public Sub Events_UpdateUsers(ByVal SlotEvent As Byte, ByVal Quotas As Integer, ByRef tUser() As Integer)

        '<EhHeader>
        On Error GoTo Events_UpdateUsers_Err

        '</EhHeader>
        Dim A           As Long

        Dim Amount      As Integer
        
        Dim MembersCant As Integer
        
        Dim CopyUsers() As tUserEvent
        
        
100     With Events(SlotEvent)
            
102         For A = LBound(tUser()) To UBound(tUser())
                Call ParticipeEvent_User(tUser(A), SlotEvent)
112         Next A

            Call Event_Initial(SlotEvent, Quotas)

        End With

        '<EhFooter>
        Exit Sub

Events_UpdateUsers_Err:
        LogError Err.description & vbCrLf & "in Events_UpdateUsers " & "at line " & Erl

        '</EhFooter>
End Sub



' # Creamos la tabla de información para la web

Public Sub Events_GenerateTableWeb()
        '<EhHeader>
        On Error GoTo Events_GenerateTableWeb_Err
        '</EhHeader>
        Dim A As Long
    
        Dim Text As String
        Dim TempObj As String
    
100     For A = 1 To MAX_EVENT_SIMULTANEO
102         With Events(A)
104             If .PrizeObj.ObjIndex > 0 Then
106                 TempObj = ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ")"
                Else
108                 TempObj = "(Ninguno)"
                End If
            
110             Text = strModality(A, .Modality) & "-" & .Inscribed & "/" & .Quotas & "-" & .LvlMin & "/" & .LvlMax & "-" & .InscriptionGld & "-" & .InscriptionEldhir & "-" & .PrizeGld & "-" & .PrizeEldhir & "-" & TempObj & vbCrLf
            End With
112     Next A
        '<EhFooter>
        Exit Sub

Events_GenerateTableWeb_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS.Events_GenerateTableWeb " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



' # Agrega daño causado guardado
Public Function Events_Add_Damage(ByVal SlotEvent As Byte, ByVal SlotEventUser As Byte, ByVal Value As Long)
    
    On Error GoTo ErrHandler
    
    With Events(SlotEvent)
        .Users(SlotEventUser).Damage = .Users(SlotEventUser).Damage + Value
        
        Call SendData(SendTarget.ToOne, .Users(SlotEventUser).ID, PrepareMessageRenderConsole("Daño total causado: " & .Users(SlotEventUser).Damage, eDamageType.d_AddMagicWord, 3000, 0))
    End With
    
    Exit Function
ErrHandler:
    
End Function


