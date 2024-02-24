Attribute VB_Name = "EventosDS"
' REFERENCIAS


'#################################
' EVENTO DE TELEPORTS
'#################################

'#################################
' EVENTO DE GRANBESTIA
'#################################


Option Explicit

Public Const MAX_EVENT_SIMULTANEO As Byte = 5

Public Const MAX_USERS_EVENT      As Byte = 64

Public Const MAX_MAP_FIGHT        As Byte = 4

Public Const MAP_TILE_VS          As Byte = 16


Public Const NPC_GRAN_BESTIA As Integer = 765

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
    
    
    Manual = 15
    
    Aracnus = 40
    HombreLobo = 50
    Minotauro = 60
    Invasion = 90
    
End Enum

Public Type tUserEvent

    ID As Integer
    Team As Byte
    Value As Integer
    Selected As Byte
    MapFight As Byte

End Type

Public Enum eFaction

    fCrim = 1
    fCiu = 2
    fLegion = 3
    fArmada = 4

End Enum

Private Type tEvents
    Name As String
    
    Enabled As Boolean
    Run As Boolean
    Modality As eModalityEvent
    TeamCant As Byte
    
    Quotas As Byte
    Inscribed As Byte
    
    LvlMax As Byte
    LvlMin As Byte
    
    GldInscription As Long
    DspInscription As Long
    
    AllowedClasses() As Byte
    AllowedFaction() As Byte
    
    PrizeAccumulated As Boolean
    PrizeDsp As Integer
    PrizeGld As Long
    PrizeObj As Obj
    
    LimitRed As Integer
    
    ValidItem As Boolean
    WinFollow As Boolean
                      
    TimeInscription As Long
    TimeCancel As Long
    TimeCount As Long
    TimeFinish As Long
    TimeInit As Long
    
    Users() As tUserEvent
    
    ' Por si alguno es con NPC
    NpcIndex As Integer
    
    ' Por si cambia el body del personaje y saca todo lo otro.
    CharBody As Integer
    CharHp As Integer
    
    npcUserIndex As Integer
    
    ChangeClass As Byte
    ChangeRaze As Byte
    InvenFree As Byte

End Type

Public Events(1 To MAX_EVENT_SIMULTANEO) As tEvents

Private Type tMap

    Run As Boolean
    Map As Integer
    X As Byte
    Y As Byte

End Type

Private Type tMapEvent

    Fight(1 To MAX_MAP_FIGHT) As tMap
    TeleportWin As tMap
End Type

Public MapEvent As tMapEvent

Public Sub LoadMapEvent()

    With MapEvent
        .Fight(1).Run = False
        .Fight(1).Map = 63
        .Fight(1).X = 16 '+17
        .Fight(1).Y = 12 '+17
              
        .Fight(2).Run = False
        .Fight(2).Map = 63
        .Fight(2).X = 16 '+17
        .Fight(2).Y = 41 '+17
        .Fight(3).Run = False
        .Fight(3).Map = 63
        .Fight(3).X = 16 '+17
        .Fight(3).Y = 68 '+17
              
        .Fight(4).Run = False
        .Fight(4).Map = 63
        .Fight(4).X = 46 '+17
        .Fight(4).Y = 12 '+17
        
        .TeleportWin.Map = 65
        .TeleportWin.X = 25
        .TeleportWin.Y = 15
    End With

End Sub

'/MANEJO DE LOS TIEMPOS '/
Public Sub LoopEvent()

10  On Error GoTo error

    Dim LoopC As Long

    Dim LoopY As Integer
          
20  For LoopC = 1 To MAX_EVENT_SIMULTANEO

30      With Events(LoopC)

40          If .Enabled Then
50              If .TimeInscription > 0 Then
60                  .TimeInscription = .TimeInscription - 1
                              
70                  Select Case .TimeInscription

                        Case Is <= 5

                            If .TimeInscription > 0 Then
                                If .TimeInscription = 1 Then
                                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & .TimeInscription & " segundo.", FontTypeNames.FONTTYPE_GUILD)
                                Else
                                    
                                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & .TimeInscription & " segundos.", FontTypeNames.FONTTYPE_GUILD)
                                End If
                            End If

                        Case 60
90                          'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(29, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInscription / 60), , , , strModality(LoopC, .Modality))
                                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInscription / 60) & " minuto.", FontTypeNames.FONTTYPE_GUILD)

100                         Case 120
110                             'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(28, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInscription / 60), , , , strModality(LoopC, .Modality))
                                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInscription / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)

120                         Case 180
130                             'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(28, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInscription / 60), , , , strModality(LoopC, .Modality))
                                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInscription / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)

140                         Case 240
150                             'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(28, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInscription / 60), , , , strModality(LoopC, .Modality))
                                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInscription / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
160                     End Select
                          
170                     If .TimeInscription <= 0 Then
180                         'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(30, FontTypeNames.FONTTYPE_GUILD, , , , , strModality(LoopC, .Modality))
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Inscripciones abiertas. /ENTRAR " & UCase$(strModality(LoopC, .Modality)) & " para ingresar al evento.", FontTypeNames.FONTTYPE_GUILD)
200                     End If
                      
210                 End If
                      
220                 If (.TimeCancel > 0) And (Not .Run) Then
230                     .TimeCancel = .TimeCancel - 1
                          
240                     If .TimeCancel <= 0 Then
                            'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Ha sido cancelado ya que no se completaron los cupos.", FontTypeNames.FONTTYPE_WARNING)
250                         EventosDS.CloseEvent LoopC, "Evento " & strModality(LoopC, .Modality) & " cancelado.", True
260                     End If
270                 End If
                      
                    If .TimeInit > 0 Then
                        .TimeInit = .TimeInit - 1
                            
                        If .TimeInit <= 0 Then
                            Call InitEvent(LoopC)
                            
                        End If
                    End If
                      
280                 If .TimeCount > 0 Then
290                     .TimeCount = .TimeCount - 1
                          
300                     For LoopY = LBound(.Users()) To UBound(.Users())

310                         If .Users(LoopY).ID > 0 Then
320                             If .TimeCount = 0 Then
                                    WriteConsoleMsg .Users(LoopY).ID, "Cuenta» ¡Comienza!", FontTypeNames.FONTTYPE_FIGHT
330                                 'WriteShortMsj .Users(LoopY).Id, 31, FontTypeNames.FONTTYPE_FIGHT
340                             Else
                                    WriteConsoleMsg .Users(LoopY).ID, "Cuenta» " & .TimeCount, FontTypeNames.FONTTYPE_GUILD
350                                 'WriteShortMsj .Users(LoopY).Id, 32, FontTypeNames.FONTTYPE_GUILD, .TimeCount
360                             End If
370                         End If

380                     Next LoopY

390                 End If
                      
400                 If .NpcIndex > 0 Then
                        If .Modality = DagaRusa Then
410                         If Events(Npclist(.NpcIndex).flags.SlotEvent).TimeCount > 0 Then Exit Sub
420                         Call DagaRusa_MoveNpc(.NpcIndex)
                        End If
430                 End If
                      
440                 If .TimeFinish > 0 Then
450                     .TimeFinish = .TimeFinish - 1
                          
460                     If .TimeFinish = 0 Then
470                         Call FinishEvent(LoopC)
480                     End If
490                 End If
500             End If
          
510         End With

520     Next LoopC
          
530     Exit Sub

error:
540     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : LoopEvent()"
End Sub

'/ FIN MANEJO DE LOS TIEMPOS
Public Function SetInfoEvento() As String

    Dim strTemp As String

    Dim LoopC   As Integer
          
10  For LoopC = 1 To EventosDS.MAX_EVENT_SIMULTANEO

20      With Events(LoopC)

30          If .Enabled Then
40              strTemp = strModality(LoopC, .Modality)
50              SetInfoEvento = SetInfoEvento & strTemp & "» " & strDescEvent(LoopC, .Modality) & ". Se ingresa mediante: /ENTRAR " & UCase$(strTemp)
                      
60              If .Run Then
70                  SetInfoEvento = SetInfoEvento & " Inscripciones cerradas."
80              Else

90                  If .TimeInscription > 0 Then
100                         SetInfoEvento = SetInfoEvento & " Inscripciones abren en " & Int(.TimeInscription / 60) & " minuto/s"
110                     Else
120                         SetInfoEvento = SetInfoEvento & " Inscripciones abiertas."
130                     End If
140                 End If
                      
150                 SetInfoEvento = SetInfoEvento & vbCrLf
160             End If

170         End With

180     Next LoopC

End Function

'// Funciones generales '//
Private Function FreeSlotEvent() As Byte

    Dim LoopC As Integer
          
10  For LoopC = 1 To MAX_EVENT_SIMULTANEO

20      If Not Events(LoopC).Enabled Then
30          FreeSlotEvent = LoopC

40          Exit For

50      End If

60  Next LoopC

End Function

Private Function Event_ModalityRepeat(ByVal Modality As eModalityEvent) As Boolean

    Dim LoopC As Integer
          
10  For LoopC = 1 To MAX_EVENT_SIMULTANEO

20      If Events(LoopC).Modality = Modality Then
30          Event_ModalityRepeat = True

40          Exit Function

50      End If

60  Next LoopC

End Function

Private Function FreeSlotUser(ByVal SlotEvent As Byte) As Byte

    Dim LoopC As Integer
          
10  With Events(SlotEvent)

20      For LoopC = 1 To MAX_USERS_EVENT

30          If .Users(LoopC).ID = 0 Then
40              FreeSlotUser = LoopC

50              Exit For

60          End If

70      Next LoopC

80  End With
          
End Function

Private Function FreeSlotArena() As Byte

    Dim LoopC As Integer
          
10  FreeSlotArena = 0
          
20  For LoopC = 1 To MAX_MAP_FIGHT

30      If MapEvent.Fight(LoopC).Run = False Then
40          FreeSlotArena = LoopC

50          Exit For

60      End If

70  Next LoopC

End Function

Public Function strUsersEvent(ByVal SlotEvent As Byte) As String

    ' Texto que marca los personajes que están en el evento.
    Dim LoopC As Integer
          
10  With Events(SlotEvent)

20      For LoopC = LBound(.Users()) To UBound(.Users())

30          If .Users(LoopC).ID > 0 Then
40              strUsersEvent = strUsersEvent & UserList(.Users(LoopC).ID).Name & "-"
            Else
                strUsersEvent = strUsersEvent & "(Vacio)" & "-"
50          End If

60      Next LoopC

70  End With

End Function

Private Function CheckAllowedClasses(ByRef AllowedClasses() As Byte) As String

    Dim LoopC As Integer
    Dim Valid As Boolean: Valid = True
    
10  For LoopC = 1 To NUMCLASES

20      If AllowedClasses(LoopC) = 1 Then
30          If CheckAllowedClasses = vbNullString Then
40              CheckAllowedClasses = ListaClases(LoopC)
50          Else
60              CheckAllowedClasses = CheckAllowedClasses & ", " & ListaClases(LoopC)
70          End If
        Else
            Valid = False
80      End If

90  Next LoopC

    If Valid Then
        CheckAllowedClasses = "TODAS"
    End If
End Function

Private Function Events_CheckAllowed_Faction(ByRef AllowedFaction() As Byte) As String

    Dim LoopC As Integer
    Dim Valid As Boolean: Valid = True
    
10  For LoopC = 1 To 4

20      If AllowedFaction(LoopC) = 1 Then
            Events_CheckAllowed_Faction = Events_CheckAllowed_Faction & Faction_String(LoopC) & ", "
        Else
            Valid = False
80      End If

90  Next LoopC

    If Len(Events_CheckAllowed_Faction) > 0 Then
         Events_CheckAllowed_Faction = Left$(Events_CheckAllowed_Faction, Len(Events_CheckAllowed_Faction) - 2)
    End If

    If Valid Then
        Events_CheckAllowed_Faction = "TODAS"
    End If
End Function

Private Function SearchLastUserEvent(ByVal SlotEvent As Byte) As Integer

    ' Busca el último usuario que está en el torneo. En todos los eventos será el ganador.
          
    Dim LoopC As Integer
          
10  With Events(SlotEvent)

20      For LoopC = LBound(.Users()) To UBound(.Users())

30          If .Users(LoopC).ID > 0 Then
40              SearchLastUserEvent = .Users(LoopC).ID

50              Exit For

60          End If

70      Next LoopC

80  End With

End Function

Private Function SearchSlotEvent(ByVal Modality As String) As Byte

    Dim LoopC As Integer
          
10  SearchSlotEvent = 0
          
20  For LoopC = 1 To MAX_EVENT_SIMULTANEO

30      With Events(LoopC)
            If .Modality = Manual Then
                If StrComp(UCase$(.Name), UCase$(Modality)) = 0 Then
                    SearchSlotEvent = LoopC
                    Exit For
                End If
            Else
40              If StrComp(UCase$(strModality(LoopC, .Modality)), UCase$(Modality)) = 0 Then
50                  SearchSlotEvent = LoopC
    
60                  Exit For
    
70              End If
            End If
80      End With

90  Next LoopC

End Function

Public Sub EventWarpUser(ByVal UserIndex As Integer, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)

    ' // NUEVO
    
10  On Error GoTo error

    ' Teletransportamos a cualquier usuario que cumpla con la regla de estar en un evento.
          
    Dim Pos As WorldPos
          
20  With UserList(UserIndex)
30      Pos.Map = Map
40      Pos.X = X
50      Pos.Y = Y
              
60      ClosestStablePos Pos, Pos
70      WarpUserChar UserIndex, Pos.Map, Pos.X, Pos.Y, False
          
80  End With
          
90  Exit Sub

error:
100     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : EventWarpUser()"
End Sub

Private Sub ResetEvent(ByVal Slot As Byte)

10  On Error GoTo error

    Dim LoopC As Integer
          
20  With Events(Slot)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
50              AbandonateEvent .Users(LoopC).ID, False
60          End If

70      Next LoopC
              
80      If .NpcIndex > 0 Then Call QuitarNPC(.NpcIndex)
              
90          .Enabled = False
100         .Run = False
110         .npcUserIndex = 0
120         .TimeFinish = 0
130         .TeamCant = 0
140         .Quotas = 0
150         .Inscribed = 0
            .Name = 0
160         .DspInscription = 0
170         .GldInscription = 0
180         .LvlMax = 0
190         .LvlMin = 0
200         .TimeCancel = 0
210         .NpcIndex = 0
220         .TimeInscription = 0
            .TimeInit = 0
230         .TimeCount = 0
240         .CharBody = 0
250         .CharHp = 0
260         .Modality = 0
            .ChangeClass = 0
            .ChangeRaze = 0
            .InvenFree = 0
            .PrizeObj.ObjIndex = 0
            .PrizeObj.Amount = 0
            .PrizeDsp = 0
            .PrizeGld = 0
              
270         For LoopC = LBound(.AllowedClasses()) To UBound(.AllowedClasses())
280             .AllowedClasses(LoopC) = 0
290         Next LoopC
              
300     End With

310     Exit Sub

error:
320     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ResetEvent()"
End Sub

Private Function CheckUserEvent(ByVal UserIndex As Integer, _
                                ByVal SlotEvent As Byte, _
                                ByRef ErrorMsg As String) As Boolean

10  On Error GoTo error

20  CheckUserEvent = False
              
30  With UserList(UserIndex)

40      If .flags.Muerto Then
50          ErrorMsg = "Estás muerto."

60          Exit Function

70      End If

80      If .flags.Mimetizado Then
90          ErrorMsg = "Estás mimetizado."

100             Exit Function

110         End If
              
120         'If .flags.Montando Then
130         'ErrorMsg = 35
140         'Exit Function
150         'End If
              
160         If .flags.Invisible Then
170             ErrorMsg = "Estás invisible."

180             Exit Function

190         End If
              
200         If .flags.SlotEvent > 0 Then
210             ErrorMsg = "Estás en Evento."

220             Exit Function

230         End If

            If .flags.Desafiando > 0 Then
                ErrorMsg = "Estás desafiando."
                
                Exit Function
            End If
              
240         If .flags.SlotReto > 0 Or .flags.SlotFast > 0 Then
250             ErrorMsg = "Estás en Reto."

260             Exit Function

270         End If
              
320         If .Counters.Pena > 0 Then
330             ErrorMsg = "Estás en la carcel pinche cabron."

340             Exit Function

350         End If
            
            If Not Is_Map_valid(UserIndex) Then
                ErrorMsg = "Estás en un mapa inválido."
                Exit Function
            End If
            
            If .flags.SlotCVC > 0 Then
                ErrorMsg = "Estás en otro enfrentamiento."
                
                Exit Function
            End If
              
360         If MapInfo(.Pos.Map).Pk Then
370             ErrorMsg = "Estás demasiado lejos de la ciudad."

380             Exit Function

390         End If
              
400         If .flags.Comerciando Then
410             ErrorMsg = "Estás comerciando buguero."

420             Exit Function

430         End If
              
440         If Not Events(SlotEvent).Enabled Or Events(SlotEvent).TimeInscription > 0 Then
450             ErrorMsg = "Las inscripciones no estan abiertas. Estate atento!"

460             Exit Function

470         End If
              
480         If Events(SlotEvent).Run Then
490             ErrorMsg = "El evento ya completo los cupos. Mejor suerte para la próxima."

500             Exit Function

510         End If
              
520         If Events(SlotEvent).LvlMin <> 0 Then
530             If Events(SlotEvent).LvlMin > .Stats.Elv Then
540                 ErrorMsg = "Tu nivel no te permite entrar al evento."

550                 Exit Function

560             End If
570         End If
              
580         If Events(SlotEvent).LvlMin <> 0 Then
590             If Events(SlotEvent).LvlMax < .Stats.Elv Then
600                 ErrorMsg = "Tu nivel no te permite entrar al evento."

610                 Exit Function

620             End If
630         End If
              
640         If Events(SlotEvent).AllowedClasses(.Clase) = 0 Then
650             ErrorMsg = "Tu clase no te permite entrar al evento."

660             Exit Function

670         End If
              
680         If Events(SlotEvent).GldInscription > .Stats.Gld Then
690             ErrorMsg = "No tienes suficientes Monedas de Oro."

700             Exit Function

710         End If
              
720         If Events(SlotEvent).DspInscription > .Stats.Eldhir Then
740             ErrorMsg = "No tienes suficientes Monedas de Eldhir."

750             Exit Function

770         End If
              
780         If Events(SlotEvent).Inscribed = Events(SlotEvent).Quotas Then
790             ErrorMsg = "El evento completo los cupos."

800             Exit Function

810         End If

            If Events(SlotEvent).InvenFree > 0 Then
                If .Invent.NroItems > 0 Then
                    ErrorMsg = "Debes tener el inventario vacío para poder participar de este evento"

                    Exit Function

                End If
            End If
            
            If Events(SlotEvent).LimitRed > 0 Then
                If TieneObjetos(POCION_ROJA, Events(SlotEvent).LimitRed + 1, UserIndex) Then
                    ErrorMsg = "El evento permite solo " & Events(SlotEvent).LimitRed & " pociones rojas."
                    
                    Exit Function
                End If
            End If
            
            
            ' NO permitimos criminales
            If Events(SlotEvent).AllowedFaction(eFaction.fCrim) = 0 Then
                If criminal(UserIndex) Then
                    ErrorMsg = "El evento no permite que ingresen criminales."
                    Exit Function
                End If
            End If
            
            ' NO permitimos ciudadanos
            If Events(SlotEvent).AllowedFaction(eFaction.fCiu) = 0 Then
                If Not criminal(UserIndex) Then
                    ErrorMsg = "El evento no permite que ingresen ciudadanos."
                    Exit Function
                End If
            End If
            
            ' NO permitimos Legionarios
            If Events(SlotEvent).AllowedFaction(eFaction.fLegion) = 0 Then
                If .Faction.Status = r_Caos Then
                    ErrorMsg = "El evento no permite que ingresen miembros de la Legión Oscura."
                    Exit Function
                End If
            End If
            
            ' NO permitimos Legionarios
            If Events(SlotEvent).AllowedFaction(eFaction.fArmada) = 0 Then
                If .Faction.Status = r_Armada Then
                    ErrorMsg = "El evento no permite que ingresen miembros de la Armada Real."
                    Exit Function
                End If
            End If
            
820     End With

830     CheckUserEvent = True
          
840     Exit Function

error:
850     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CheckUserEvent()"
End Function

Public Function SendInfoEvent(ByVal UserIndex As Integer, ByVal Slot As Byte) As Byte

    Dim Count As Byte
    Dim TempStr As String
    
    On Error GoTo ErrHandler

    With Events(Slot)

        If .Enabled Then
            
            TempStr = "Evento " & strModality(Slot, .Modality) & "» Cupos disponibles: " & .Inscribed & "/" & .Quotas & vbCrLf
            TempStr = TempStr & "Nivel permitido: " & .LvlMin & "/" & .LvlMax & ". Oro requerido: " & .GldInscription & ". Eldhires requeridos: " & .DspInscription & vbCrLf
            TempStr = TempStr & "Premio: " & .PrizeGld & " Monedas de Oro y " & .PrizeDsp & " Monedas Eldhir"
            
            If .PrizeObj.ObjIndex > 0 Then
                TempStr = TempStr & " + ¡Premio Extra! " & ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ")" & vbCrLf
            Else
                TempStr = TempStr & vbCrLf
            End If
            
            TempStr = TempStr & "Tipea '/ENTRAR " & UCase$(strModality(Slot, .Modality)) & "' para ingresar al evento. "
            
            
            Call WriteConsoleMsg(UserIndex, TempStr, FontTypeNames.FONTTYPE_INFOGREEN)
            
            If .TimeInscription = 0 Then
                If .Run Then
                    Call WriteConsoleMsg(UserIndex, "Las inscripciones están cerradas.", FontTypeNames.fonttype_infored)
                Else
                    Call WriteConsoleMsg(UserIndex, "Las inscripciones están abiertas.", FontTypeNames.FONTTYPE_INFOGREEN)
                End If

            ElseIf .TimeInscription <= 60 Then
                Call WriteConsoleMsg(UserIndex, "Las inscripciones abren en " & .TimeInscription & " segundos.", FontTypeNames.FONTTYPE_INFOGREEN)
            Else
                Call WriteConsoleMsg(UserIndex, "Las inscripciones abren en " & Int(.TimeInscription / 60) & " minuto.", FontTypeNames.fonttype_infored)
            End If
            
            SendInfoEvent = 1
        End If

    End With
    
    Exit Function

ErrHandler:
    Call LogError("Error en InfoEvent")
End Function

' EDICIÓN GENERAL
Public Function strModality(ByVal SlotEvent As Byte, _
                            ByVal Modality As eModalityEvent) As String

    ' Modalidad de cada evento
          
10  Select Case Modality

        Case eModalityEvent.CastleMode
20          strModality = "REYvsREY"
                  
30      Case eModalityEvent.DagaRusa
40          strModality = "DagaRusa"
                  
50      Case eModalityEvent.DeathMatch
60          strModality = "DeathMatch"
                  
70      Case eModalityEvent.Aracnus
80          strModality = "Aracnus"
                  
90      Case eModalityEvent.HombreLobo
100         strModality = "HombreLobo"
                  
110     Case eModalityEvent.Minotauro
120         strModality = "Minotauro"
              
130    Case eModalityEvent.Busqueda
140             strModality = "Busqueda"
              
150     Case eModalityEvent.Unstoppable
160             strModality = "Imparable"
              
        Case eModalityEvent.JuegosDelHambre
            strModality = "JDH"
            
170     Case eModalityEvent.Invasion
180             strModality = "Invasion"
                  
190     Case eModalityEvent.Enfrentamientos
200          strModality = Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant

        Case eModalityEvent.Manual
            strModality = Events(SlotEvent).Name
            
        Case eModalityEvent.Teleports
            strModality = "Teleports"
            
        Case eModalityEvent.GranBestia
            strModality = "GranBestia"
            
210     End Select

End Function

Private Function strDescEvent(ByVal SlotEvent As Byte, _
                              ByVal Modality As eModalityEvent) As String

    ' Descripción del evento en curso.
10  Select Case Modality

        Case eModalityEvent.CastleMode
20          strDescEvent = "» Los usuarios entrarán de forma aleatorea para formar dos equipos. Ambos equipos deberán defender a su rey y a su vez atacar al del equipo contrario."

30      Case eModalityEvent.DagaRusa
40          strDescEvent = "» Los usuarios se teletransportarán a una posición donde estará un asesino dispuesto a apuñalarlos y acabar con su vida. El último que quede en pie es el ganador del evento."

50      Case eModalityEvent.DeathMatch
60          strDescEvent = "» Los usuarios ingresan y luchan en una arena donde se toparan con todos los demás concursantes. El que logre quedar en pie, será el ganador."

70      Case eModalityEvent.Aracnus
80          strDescEvent = "» Un personaje es escogido al azar, para convertirse en una araña gigante la cual podrá envenenar a los demas concursantes acabando con su vida en el evento."

90      Case eModalityEvent.Busqueda
100             strDescEvent = "» Los personajes son teletransportados en un mapa donde su función principal será la recolección de objetos en el piso, para que así luego de tres minutos, el que recolecte más, ganará el evento."

110         Case eModalityEvent.Unstoppable
120             strDescEvent = "» Los personajes lucharan en un TODOS vs TODOS, donde los muertos no irán a su mapa de origen, si no que volverán a revivir para tener chances de ganar el evento. El que logre matar más personajes, ganará el evento."
            
            Case eModalityEvent.JuegosDelHambre
                strDescEvent = "Descripcion"
                
130         Case eModalityEvent.Invasion
140             strDescEvent = "» Los personajes son llevados a un mapa donde aparecerán criaturas únicas de DesteriumAO, cada criatura dará una recompensa única y los usuarios tendrán chances de entrenar sus personajes."

150         Case eModalityEvent.Enfrentamientos

160             If Events(SlotEvent).TeamCant = 1 Then
170                 strDescEvent = "» Los usuarios combatirán en duelos 1vs1"
180             Else
190                 strDescEvent = "» Los usuarios combatirán en duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & " donde se escogerán las parejas al azar."
200             End If

210     End Select

End Function

Private Sub InitEvent(ByVal SlotEvent As Byte)
    
10  Select Case Events(SlotEvent).Modality

        Case eModalityEvent.CastleMode
20          Call InitCastleMode(SlotEvent)
                  
30      Case eModalityEvent.DagaRusa
40          Call InitDagaRusa(SlotEvent)
                  
50      Case eModalityEvent.DeathMatch
60          Call InitDeathMatch(SlotEvent)
                  
70      Case eModalityEvent.Aracnus
80          Call InitEventTransformation(SlotEvent, 254, 6500, 60, 70, 36)
                  
90      Case eModalityEvent.HombreLobo
100             Call InitEventTransformation(SlotEvent, 255, 3500, 60, 70, 36)
                  
110         Case eModalityEvent.Minotauro
120             Call InitEventTransformation(SlotEvent, 253, 2500, 60, 70, 36)
              
130         Case eModalityEvent.Busqueda
140             Call InitBusqueda(SlotEvent)
                  
150         Case eModalityEvent.Unstoppable
160             InitUnstoppable SlotEvent
            
            Case eModalityEvent.JuegosDelHambre
                JDH_Init SlotEvent
              
180         Case eModalityEvent.Enfrentamientos
190             Call InitFights(SlotEvent)
            
            Case eModalityEvent.Teleports
                Call Events_Teleports_Init(SlotEvent)
            
            Case eModalityEvent.GranBestia
                Call Events_GranBestia_Init(SlotEvent)
                
200         Case Else

210             Exit Sub
              
220     End Select

230     Exit Sub

error:
240     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitEvent() EN EL EVENTO " & Events(SlotEvent).Modality & "."
End Sub

Public Function CanAttackUserEvent(ByVal UserIndex As Integer, _
                                   ByVal Victima As Integer) As Boolean
          
    ' Si el personaje es del mismo team, no se puede atacar al usuario.
    Dim VictimaSlotUserEvent As Byte
          
10  VictimaSlotUserEvent = UserList(Victima).flags.SlotUserEvent
          
    If UserList(UserIndex).flags.SlotEvent > 0 And UserList(Victima).flags.SlotEvent > 0 Then

        With UserList(UserIndex)

40          If Events(.flags.SlotEvent).Users(VictimaSlotUserEvent).Team = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
50              CanAttackUserEvent = False

60              Exit Function

70          End If

            End With

        End If
   
        CanAttackUserEvent = True
          
110     Exit Function

error:
120     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAttackUserEvent()"
End Function

Private Sub PrizeUser(ByVal UserIndex As Integer, ByRef strReward As String)

10  On Error GoTo error
          
    ' Premios de los eventos
          
    Dim SlotEvent     As Byte

    Dim SlotUserEvent As Byte

    Dim Obj           As Obj

20  SlotEvent = UserList(UserIndex).flags.SlotEvent
30  SlotUserEvent = UserList(UserIndex).flags.SlotUserEvent
          
40  With Events(SlotEvent)
             
            If .PrizeGld = 0 And .PrizeDsp = 0 And .PrizeObj.ObjIndex = 0 Then
                strReward = "Premio entregado: NINGUNO"
                
            Else
                strReward = "Premio entregado: "
                
                If .PrizeGld > 0 Then
                    UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + .PrizeGld
                    Call WriteUpdateGold(UserIndex)
                       
                    strReward = strReward & .PrizeGld & " Monedas de Oro. "
                End If
                    
                If .PrizeDsp > 0 Then
                    UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir + .PrizeDsp
                    Call WriteUpdateDsp(UserIndex)
                       
                    strReward = strReward & .PrizeDsp & " Monedas de Eldhir. "
                End If
                    
                If .PrizeObj.ObjIndex > 0 Then
                    If Not MeterItemEnInventario(UserIndex, .PrizeObj) Then
                        WriteConsoleMsg UserIndex, "Tu premio OBJETO no ha sido entregado, envia esta foto a un Game Master.", FontTypeNames.FONTTYPE_INFO
                        LogEventos ("Personaje " & UserList(UserIndex).Name & " no recibió: " & .PrizeObj.ObjIndex & " (x" & .PrizeObj.Amount & ")")
                    End If
                       
                    strReward = strReward & ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ")"
                End If
            End If
              
230         With UserList(UserIndex)
240             .Stats.TorneosGanados = .Stats.TorneosGanados + 1
                
                Call RankUser_AddPoint(UserIndex, 1)
250         End With
              
260
270     End With
          
280     Exit Sub

error:
290     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : PrizeUser()"
End Sub

Private Sub ChangeBodyEvent(ByVal SlotEvent As Byte, _
                            ByVal UserIndex As Integer, _
                            ByVal ChangeHead As Boolean)

10  On Error GoTo error

    ' En caso de que el evento cambie el body, de lo cambiamos.
20  With UserList(UserIndex)
30      .CharMimetizado.Body = .Char.Body
40      .CharMimetizado.Head = .Char.Head
50      .CharMimetizado.CascoAnim = .Char.CascoAnim
60      .CharMimetizado.ShieldAnim = .Char.ShieldAnim
70      .CharMimetizado.WeaponAnim = .Char.WeaponAnim

80      .Char.Body = Events(SlotEvent).CharBody
90      .Char.Head = IIf(ChangeHead = False, .Char.Head, 0)
100         .Char.CascoAnim = 0
110         .Char.ShieldAnim = 0
120         .Char.WeaponAnim = 0
                      
130         'ChangeUserChar UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim
140         RefreshCharStatus UserIndex
          
150     End With
          
160     Exit Sub

error:
170     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ChangeBodyEvent()"
End Sub

Private Function ResetBodyEvent(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

10  On Error GoTo error

    ' En caso de que el evento cambie el body del personaje, se lo restauramos.
          
20  With UserList(UserIndex)

30      If .flags.Muerto Then Exit Function
        'If Events(SlotEvent).Users(.flags.SlotUserEvent).Selected = 0 Then Exit Function
              
40      If .CharMimetizado.Body > 0 Then
50          .Char.Body = .CharMimetizado.Body
60          .Char.Head = .CharMimetizado.Head
70          .Char.CascoAnim = .CharMimetizado.CascoAnim
80          .Char.ShieldAnim = .CharMimetizado.ShieldAnim
90          .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                  
100             .CharMimetizado.Body = 0
110             .CharMimetizado.Head = 0
120             .CharMimetizado.CascoAnim = 0
130             .CharMimetizado.ShieldAnim = 0
140             .CharMimetizado.WeaponAnim = 0
                .CharMimetizado.AuraIndex = 0
                  
150             .ShowName = True
                  
160             ChangeUserChar UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex
170             RefreshCharStatus UserIndex
180         End If
          
190     End With
          
200     Exit Function

error:
210     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ResetBodyEvent()"
End Function

Private Sub ChangeHpEvent(ByVal UserIndex As Integer)

10  On Error GoTo error

    ' En caso de que el evento edite la vida del personaje, se la editamos.
          
    Dim SlotEvent As Byte
          
20  With UserList(UserIndex)
30      SlotEvent = .flags.SlotEvent
              
40      .Stats.OldHp = .Stats.MaxHp
        
50      .Stats.MaxHp = Events(SlotEvent).CharHp
60      .Stats.MinHp = .Stats.MaxHp
              
70      WriteUpdateUserStats UserIndex
          
80  End With

90  Exit Sub

error:
100     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ChangeHpEvent()"
End Sub

Private Sub ResetHpEvent(ByVal UserIndex As Integer)

10  On Error GoTo error

    ' En caso de que el evento haya editado la vida de un personaje, se la volvemos a restaurar.
          
20  With UserList(UserIndex)

30      If .Stats.OldHp = 0 Then Exit Sub
40      .Stats.MaxHp = .Stats.OldHp
        '.Stats.MinHp = .Stats.MaxHp
50      .Stats.OldHp = 0
60      WriteUpdateHP UserIndex
              
70  End With
          
80  Exit Sub

error:
90  LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ResetHpEvent()"
End Sub

'// Fin Funciones generales '//

Public Function NewEvent(ByVal Modality As eModalityEvent, _
                         ByVal Name As String, _
                         ByVal Quotas As Byte, _
                         ByVal LvlMin As Byte, _
                         ByVal LvlMax As Byte, _
                         ByVal GldInscription As Long, _
                         ByVal DspInscription As Long, _
                         ByVal TimeInscription As Long, _
                         ByVal TimeCancel As Long, _
                         ByVal TeamCant As Byte, _
                         ByVal PrizeAccumulated As Boolean, _
                         ByVal LimitRed As Integer, _
                         ByVal PrizeDsp As Integer, _
                         ByVal PrizeGld As Long, _
                         ByVal ObjIndex As Integer, _
                         ByVal ObjAmount As Integer, _
                         ByVal WinFollow As Boolean, _
                         ByVal ValidItem As Boolean, _
                         ByVal InvFree As Byte, _
                         ByVal ChangeClass As eClass, _
                         ByVal ChangeRaze As eRaza, _
                         ByRef AllowedFaction() As Byte, _
                         ByRef AllowedClasses() As Byte) As Boolean
                          
10  On Error GoTo error
                          
    Dim Slot    As Integer

    Dim strTemp As String

20  Slot = FreeSlotEvent()
          
    If Event_ModalityRepeat(Modality) Then
        NewEvent = False

        'Call WriteConsoleMsg(UserIndex, "Ya existe un evento con esa modalidad. ¡Mejor cuidado para la próxima!", FontTypeNames.FONTTYPE_INFO)
        Exit Function

    End If
          
30  If Slot = 0 Then
        NewEvent = False

        'WriteConsoleMsg UserIndex, "No hay más lugar disponible para crear un evento simultaneo. Espera a que termine alguno o bien cancela alguno.", FontTypeNames.FONTTYPE_INFO
40      'WriteShortMsj UserIndex, 48, FontTypeNames.FONTTYPE_INFO
50      Exit Function

60  Else

70      With Events(Slot)
80              .Enabled = True
                .Name = Name
90              .Modality = Modality
100             .TeamCant = TeamCant
110             .Quotas = Quotas
120             .LvlMin = LvlMin
130             .LvlMax = LvlMax
140             .GldInscription = GldInscription
150             .DspInscription = DspInscription
160             .AllowedClasses = AllowedClasses
                .AllowedFaction = AllowedFaction
170             .TimeInscription = TimeInscription
180             .TimeCancel = TimeCancel
                
                .ValidItem = ValidItem
                .PrizeAccumulated = PrizeAccumulated
                .LimitRed = LimitRed
                .PrizeDsp = PrizeDsp
                .PrizeGld = PrizeGld
                .PrizeObj.ObjIndex = ObjIndex
                .PrizeObj.Amount = ObjAmount
                .WinFollow = WinFollow
                .ChangeClass = ChangeClass
                .ChangeRaze = ChangeRaze
                .InvenFree = InvFree
                  
190             ReDim .Users(1 To .Quotas) As tUserEvent
                  
                strTemp = "Evento automático» " & strModality(Slot, .Modality) & vbCrLf
                strTemp = strTemp & "Cupos disponibles: " & .Quotas & ". Nivel requerido: " & .LvlMin & " a " & .LvlMax & IIf((.LimitRed > 0), ". Limite de rojas: " & .LimitRed, ".")
                
                Dim TextClass As String: TextClass = CheckAllowedClasses(.AllowedClasses)
                If TextClass <> "TODAS" Then
                    strTemp = strTemp & vbCrLf & "Clases permitidas: " & TextClass
                End If

                
                Dim TextFaction As String: TextFaction = Events_CheckAllowed_Faction(.AllowedFaction)
                If TextFaction <> "TODAS" Then
                    strTemp = strTemp & vbCrLf & "Facciones permitidas: " & TextFaction
                End If
                
                If .GldInscription > 0 Then
                    strTemp = strTemp & vbCrLf & "Oro requerido: " & .GldInscription & "."
                End If
                  
                If .DspInscription > 0 Then
                    strTemp = strTemp & vbCrLf & "Eldhires requerido: " & .DspInscription & "."
                End If
                
                If .PrizeGld > 0 Or .PrizeDsp > 0 Or .PrizeObj.ObjIndex > 0 Then
                    strTemp = strTemp & vbCrLf
                End If
                
                If .PrizeGld > 0 Then strTemp = strTemp & "Premio en Oro: " & .PrizeGld & ". "
                If .PrizeDsp > 0 Then strTemp = strTemp & "Premio en Eldhires: " & .PrizeDsp & ". "
                If .PrizeObj.ObjIndex > 0 Then strTemp = strTemp & "Premio Extra: " & ObjData(.PrizeObj.ObjIndex).Name & " (x" & .PrizeObj.Amount & ") "
                
                If .ChangeClass > 0 Then
                    strTemp = strTemp & vbCrLf & "¡¡Atención!! El evento tiene la modalidad de cambio de clase. Pasaras a ser " & ListaClases(.ChangeClass)
                End If
                
                If .ChangeRaze > 0 Then
                    strTemp = strTemp & vbCrLf & "¡¡Atención!! El evento tiene la modalidad de cambio de raza. Pasaras a ser " & ListaRazas(.ChangeRaze)
                End If
                  
                If .InvenFree > 0 Then
                    strTemp = strTemp & vbCrLf & "¡¡Atención!! El evento solicita tener el inventario vacío."
                End If
                  
                strTemp = strTemp & vbCrLf & "El comando para participar del evento es '/ENTRAR " & UCase$(strModality(Slot, .Modality)) & "'"
                
310             If .TimeInscription <= 60 Then
320                 strTemp = strTemp & vbCrLf & "Las inscripciones abren en " & .TimeInscription & " segundos. "
330             Else
340                 strTemp = strTemp & vbCrLf & "Las inscripciones abren en " & Int(.TimeInscription / 60) & " minutos. "
350             End If
                
                

370         End With
              
380         SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_EVENT)
390     End If
          
        NewEvent = True

400     Exit Function

error:
410     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : NewEvent()"
End Function

Private Sub GiveBack_Inscription(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim LoopC As Integer

    Dim Obj   As Obj
          
20  With Events(SlotEvent)
          
30      Obj.ObjIndex = 880
40      Obj.Amount = .DspInscription
              
50      For LoopC = LBound(.Users()) To UBound(.Users())

60          If .Users(LoopC).ID > 0 Then
70              If .DspInscription > 0 Then
80                  UserList(.Users(LoopC).ID).Stats.Eldhir = UserList(.Users(LoopC).ID).Stats.Eldhir + .DspInscription
                        WriteUpdateDsp (.Users(LoopC).ID)
                    End If
                      
130                 If .GldInscription > 0 Then
140                     UserList(.Users(LoopC).ID).Stats.Gld = UserList(.Users(LoopC).ID).Stats.Gld + .GldInscription
150                     WriteUpdateGold (.Users(LoopC).ID)
160                 End If
170             End If

180         Next LoopC

190     End With
          
200     Exit Sub

error:
210     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : GiveBack_Inscription()"
End Sub

Public Sub CloseEvent(ByVal Slot As Byte, _
                      Optional ByVal MsgConsole As String = vbNullString, _
                      Optional ByVal Cancel As Boolean = False)

10  On Error GoTo error
          
20  With Events(Slot)

        ' Devolvemos la inscripción
30      If Cancel Then
40          Call GiveBack_Inscription(Slot)
50      End If
              
60      If MsgConsole <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(MsgConsole, FontTypeNames.FONTTYPE_INFOBOLD)
              
70      Call ResetEvent(Slot)
80  End With
          
90  Exit Sub

error:
100     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CloseEvent()"
End Sub

Public Sub ParticipeEvent(ByVal UserIndex As Integer, ByVal Modality As String)

10  On Error GoTo error

    Dim ErrorMsg  As String

    Dim SlotUser  As Byte

    Dim Pos       As WorldPos

    Dim SlotEvent As Integer
          
20  SlotEvent = SearchSlotEvent(Modality)
          
30  If SlotEvent = 0 Then

        'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Error Fatal TESTEO", FontTypeNames.FONTTYPE_ADMIN)
40      Exit Sub

50  End If
          
60  With UserList(UserIndex)

70      If CheckUserEvent(UserIndex, SlotEvent, ErrorMsg) Then
80          SlotUser = FreeSlotUser(SlotEvent)
                  
90          .flags.SlotEvent = SlotEvent
100             .flags.SlotUserEvent = SlotUser
                  
110             .PosAnt.Map = .Pos.Map
120             .PosAnt.X = .Pos.X
130             .PosAnt.Y = .Pos.Y
                  
140             .Stats.Gld = .Stats.Gld - Events(SlotEvent).GldInscription
                .Stats.Eldhir = .Stats.Eldhir - Events(SlotEvent).DspInscription
                  
150             Call WriteUpdateGold(UserIndex)
                Call WriteUpdateDsp(UserIndex)
                  
170             With Events(SlotEvent)
180                 Pos.Map = 60
190                 Pos.X = 30
200                 Pos.Y = 21
                      
210                 Call FindLegalPos(UserIndex, Pos.Map, Pos.X, Pos.Y)
220                 Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
                  
230                 .Users(SlotUser).ID = UserIndex
240                 .Inscribed = .Inscribed + 1
                      
                    WriteConsoleMsg UserIndex, "Has ingresado al evento " & strModality(SlotEvent, .Modality) & ". Espera a que se completen los cupos para que comience.", FontTypeNames.FONTTYPE_INFO
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» El personaje " & UserList(UserIndex).Name & " ha ingresado al evento", FontTypeNames.FONTTYPE_INFOGREEN)
                      
                    'WriteShortMsj UserIndex, 51, FontTypeNames.FONTTYPE_INFO, , , , , strModality(SlotEvent, .Modality)
                    LogEventos "El personaje " & UserList(UserIndex).Name & " ingresó el evento de modalidad " & strModality(SlotEvent, .Modality)
                      
260                 If .Inscribed = .Quotas Then
                        .TimeCancel = 0
                        
                        ' Avisamos quienes son los personajes que ingresaron
                        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» ¡Cupos alcanzados!. Personajes inscriptos: " & Event_Text_Users(SlotEvent) & vbCrLf & "Mezclando personajes...", FontTypeNames.FONTTYPE_GUILD)
                        
                        If .Modality <> Manual Then
                            ' Mezclamos los personajes ingresados
                            Call Event_RandomUsers_Array(SlotEvent, .Users)
                        End If
                        
                        If .Modality = CastleMode Or .Modality = DeathMatch Then
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» " & Event_Text_Users_VS(SlotEvent), FontTypeNames.FONTTYPE_GUILD)
                        End If
                        
                        LogEventos "Cupos alcanzados."
280
290                     'InitEvent SlotEvent
                          
                        If .ChangeClass > 0 Then
                            .TimeInit = 180
                        Else
                            .TimeInit = 15
                        End If
                            
                        .Run = True
                          
                        Call Event_CheckModalityAndApplyEffects(SlotEvent)
        
300                     Exit Sub

310                 End If

320             End With
              
330         Else
340             WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING
              
350         End If

360     End With

370     Exit Sub

error:
380     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ParticipeEvent()"
End Sub

Public Sub Event_ClassOld(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

    With Events(SlotEvent)

        Dim A As Long

        With UserList(UserIndex)

            If .flags.TomoPocion Then
                LogEventos ("El personaje " & .Name & " era: " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & ". Atributos: " & .Stats.UserAtributosBackUP(1) & "-" & .Stats.UserAtributosBackUP(2) & "-" & .Stats.UserAtributosBackUP(3) & "-" & .Stats.UserAtributosBackUP(4) & "-" & .Stats.UserAtributosBackUP(5) & ". Oro: " & .Stats.Gld & ", D.Azules: " & .Stats.Eldhir & ". Vida: " & .Stats.MaxHp & " y maná " & .Stats.MaxMan & ". HIT: " & .Stats.MaxHit)
            Else
                LogEventos ("El personaje " & .Name & " era: " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & ". Atributos: " & .Stats.UserAtributos(1) & "-" & .Stats.UserAtributos(2) & "-" & .Stats.UserAtributos(3) & "-" & .Stats.UserAtributos(4) & "-" & .Stats.UserAtributos(5) & ". Oro: " & .Stats.Gld & ", D.Azules: " & .Stats.Eldhir & ". Vida: " & .Stats.MaxHp & " y maná " & .Stats.MaxMan & ". HIT: " & .Stats.MaxHit)
            End If

        End With
    
    End With

End Sub

Public Sub Event_CheckModalityAndApplyEffects(ByVal SlotEvent As Byte)

    With Events(SlotEvent)
    
        ' Clases y razas alteradas
        ' Cambiamos los stats de los pjs
        ' Llevamos a los pjs a la tienda
        If .ChangeClass > 0 Then
            Call Event_Change(SlotEvent)
        End If
        
    End With

End Sub

Public Sub AbandonateEvent(ByVal UserIndex As Integer, _
                           Optional ByVal MsgAbandonate As Boolean = False, _
                           Optional ByVal Forzado As Boolean = False, _
                           Optional ByVal ResetTotal As Boolean = False)
          
10  On Error GoTo error

    Dim Pos           As WorldPos

    Dim SlotEvent     As Byte

    Dim SlotUserEvent As Byte

    Dim UserTeam      As Byte

    Dim UserMapFight  As Byte
          
20  With UserList(UserIndex)
30      SlotEvent = .flags.SlotEvent
40      SlotUserEvent = .flags.SlotUserEvent
              
50      If SlotEvent > 0 And SlotUserEvent > 0 Then

60          With Events(SlotEvent)
                        
                LogEventos "El personaje " & UserList(UserIndex).Name & " abandonó el evento de modalidad " & strModality(SlotEvent, .Modality)
70
                If .Inscribed > 0 Then .Inscribed = .Inscribed - 1
                        
80              UserTeam = .Users(SlotUserEvent).Team
90              UserMapFight = .Users(SlotUserEvent).MapFight
                          
100                 .Users(SlotUserEvent).ID = 0
110                 .Users(SlotUserEvent).Team = 0
                    .Users(SlotUserEvent).Value = 0
130                 .Users(SlotUserEvent).Selected = 0
140                 .Users(SlotUserEvent).MapFight = 0
                          
150                 UserList(UserIndex).flags.SlotEvent = 0
160                 UserList(UserIndex).flags.SlotUserEvent = 0
170                 UserList(UserIndex).flags.FightTeam = 0
                                                    
                    If .ChangeClass > 0 Or .ChangeRaze > 1 Then
                        Call Event_UserResetClass(UserIndex)
                    End If
                          
180                 Select Case .Modality
                                  
240                     Case eModalityEvent.DagaRusa

                            If .Run Then
                                Call WriteUserInEvent(UserIndex)
                                  
250                             If Forzado Then
270                                 If .Users(SlotUserEvent).Value = 0 Then
280                                     Npclist(.NpcIndex).flags.InscribedPrevio = Npclist(.NpcIndex).flags.InscribedPrevio - 1
290                                 End If
300                             End If
                            End If
                                  
310                     Case eModalityEvent.Enfrentamientos

320                         If Forzado Then
330                             If UserMapFight > 0 Then
340                                 If Not Fight_CheckContinue(UserIndex, SlotEvent, UserTeam) Then
350                                     Fight_WinForzado UserIndex, SlotEvent, UserMapFight
360                                 End If
370                             End If
380                         End If
                                  
390                         If UserList(UserIndex).Counters.TimeFight > 0 Then
400                             UserList(UserIndex).Counters.TimeFight = 0
410                             Call WriteUserInEvent(UserIndex)
420                         End If

                        Case eModalityEvent.DeathMatch
                            UserList(UserIndex).flags.Mimetizado = 0
                                  
430                 End Select
                                  
                    If .Run Then
                        UserList(UserIndex).Stats.TorneosJugados = UserList(UserIndex).Stats.TorneosJugados + 1
                    End If
                          
440                 Pos.Map = UserList(UserIndex).PosAnt.Map
450                 Pos.X = UserList(UserIndex).PosAnt.X
460                 Pos.Y = UserList(UserIndex).PosAnt.Y
                          
470                 Call FindLegalPos(UserIndex, Pos.Map, Pos.X, Pos.Y)
480                 Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
                          
490                 If Events(SlotEvent).CharBody <> 0 Then
500                     Call ResetBodyEvent(SlotEvent, UserIndex)
510                 End If
                  
520                 If UserList(UserIndex).Stats.OldHp <> 0 Then
530                     ResetHpEvent UserIndex
540                 End If
                  
550                 UserList(UserIndex).ShowName = True
560                 RefreshCharStatus UserIndex
                          
                    If MsgAbandonate Then WriteConsoleMsg UserIndex, "Has abandonado el evento. Podrás recibir una pena por hacer esto.", FontTypeNames.FONTTYPE_WARNING
570                 'If MsgAbandonate Then WriteShortMsj UserIndex, 53, FontTypeNames.FONTTYPE_WARNING
                          
                    ' Abandono general del evento
                    If .Run Then
580                     If .Inscribed = 1 And Forzado Then
590                         'Call FinishEvent(SlotEvent)
                              
600                         CloseEvent SlotEvent

610                         Exit Sub

620                     End If
                    End If
                          
630             End With

640         End If
              
650     End With
          
660     Exit Sub

error:
670     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AbandonateEvent()"
End Sub


' Eventos que necesiten terminar despues de un cierto tiempo.
Private Sub FinishEvent(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim UserIndex  As Integer

    Dim IsSelected As Boolean

    Dim strReward  As String
          
20  With Events(SlotEvent)

30      Select Case .Modality
                      
150             Case eModalityEvent.Busqueda
160                 Busqueda_SearchWin SlotEvent
                      
170             Case eModalityEvent.Unstoppable
180                 Unstoppable_UserWin SlotEvent
                    
190         End Select

200     End With
          
210     Exit Sub

error:
220     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : FinishEvent()"
End Sub

'#################EVENTO CASTLE MODE##########################
Public Function CanAttackReyCastle(ByVal UserIndex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean

10  With UserList(UserIndex)

20      If .flags.SlotEvent > 0 Then
30          If Npclist(NpcIndex).flags.TeamEvent = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
40              CanAttackReyCastle = False

50              Exit Function

60          End If
70      End If
          
80      CanAttackReyCastle = True
90  End With

End Function

Private Sub CastleMode_InitRey()

10  On Error GoTo error
          
    Dim NpcIndex As Integer

    Const NumRey As Integer = 697

    Dim Pos      As WorldPos

    Dim LoopX    As Integer, LoopY As Integer

    Const Rango  As Byte = 5
          
20  For LoopX = YMinMapSize To YMaxMapSize
30      For LoopY = XMinMapSize To XMaxMapSize

40          If InMapBounds(61, LoopX, LoopY) Then
50              If MapData(61, LoopX, LoopY).NpcIndex > 0 Then
60                  Call QuitarNPC(MapData(61, LoopX, LoopY).NpcIndex)
70              End If
80          End If

90      Next LoopY
100     Next LoopX
          
        Pos.Map = 212
                  
        Pos.X = 50
        Pos.Y = 23
        NpcIndex = SpawnNpc(NumRey, Pos, False, False)
        Npclist(NpcIndex).flags.TeamEvent = 2
              
        Pos.X = 50
        Pos.Y = 80
        NpcIndex = SpawnNpc(NumRey, Pos, False, False)
        Npclist(NpcIndex).flags.TeamEvent = 1
          
200     Exit Sub

error:
210     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CastleMode_InitRey()"
End Sub

Public Sub InitCastleMode(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim LoopC    As Integer

    Dim NpcIndex As Integer

    Dim Pos      As WorldPos
          
    ' Spawn the npc castle mode
20  CastleMode_InitRey
          
30  With Events(SlotEvent)

40      For LoopC = LBound(.Users()) To UBound(.Users())

50          If .Users(LoopC).ID > 0 Then
60              If LoopC > (UBound(.Users()) / 2) Then
70                  .Users(LoopC).Team = 2
                    UserList(.Users(LoopC).ID).flags.FightTeam = 2
80                  Pos.Map = 212
90                  Pos.X = 50
100                 Pos.Y = 23
                          
110                     Call FindLegalPos(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y)
120                     Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, False)
130                 Else
140                     .Users(LoopC).Team = 1
                        UserList(.Users(LoopC).ID).flags.FightTeam = 1
150                     Pos.Map = 212
160                     Pos.X = 50
170                     Pos.Y = 80
                          
180                     Call FindLegalPos(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y)
190                     Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, False)
                          
200                 End If
210             End If

220         Next LoopC

230     End With
          
240     Exit Sub

error:
250     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitCastleMode()"
End Sub

Public Sub CastleMode_UserRevive(ByVal UserIndex As Integer)

10  On Error GoTo error

    Dim LoopC As Integer

    Dim Pos   As WorldPos
          
20  With UserList(UserIndex)

30      If .flags.SlotEvent > 0 Then
40          Call RevivirUsuario(UserIndex)
                  
50          Pos.Map = 212
            Pos.X = RandomNumber(20, 80)
            Pos.Y = RandomNumber(20, 80)
                  
80          Call ClosestLegalPos(Pos, Pos)
            'Call FindLegalPos(Userindex, Pos.Map, Pos.X, Pos.Y)
90          Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, True)
              
100         End If

110     End With
          
120     Exit Sub

error:
130     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CastleMode_UserRevive()"
End Sub

Public Sub FinishCastleMode(ByVal SlotEvent As Byte, ByVal UserEventSlot As Integer)

10  On Error GoTo error

    Dim LoopC     As Integer

    Dim strTemp   As String

    Dim NpcIndex  As Integer

    Dim MiObj     As Obj

    Dim strReward As String
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
50              If .Users(LoopC).Team = .Users(UserEventSlot).Team Then
60                  If LoopC = UserEventSlot Then
70                      CastleMode_Premio .Users(LoopC).ID, True
80                  Else
90                      CastleMode_Premio .Users(LoopC).ID, False
100                     End If
                          
110                     If strTemp = vbNullString Then
120                         strTemp = UserList(.Users(LoopC).ID).Name
130                     Else
140                         strTemp = strTemp & ", " & UserList(.Users(LoopC).ID).Name
150                     End If
                          
                        Call PrizeUser(.Users(LoopC).ID, strReward)

160                 End If
170             End If

180         Next LoopC
              
190         CloseEvent SlotEvent, "CastleMode» Ha finalizado. Ha ganado el equipo de " & UCase$(strTemp) & ". " & strReward
200     End With
          
210     Exit Sub

error:
220     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : FinishCastleMode()"
End Sub

Private Sub CastleMode_Premio(ByVal UserIndex As Integer, ByVal KillRey As Boolean)

10  On Error GoTo error

    ' Entregamos el premio del CastleMode
    Dim MiObj As Obj
          
20  With UserList(UserIndex)
30      .Stats.Gld = .Stats.Gld + 250000
        WriteConsoleMsg UserIndex, "Felicitaciones, has recibido 250.000 monedas de oro por haber ganado el evento!", FontTypeNames.FONTTYPE_INFO
40    ' WriteShortMsj UserIndex, 54, FontTypeNames.FONTTYPE_INFO, , , , 250000
              
50      If KillRey Then
            WriteConsoleMsg UserIndex, "Hemos notado que has aniquilado con la vida del rey oponente. ¡FELICITACIONES! Aquí tienes tu recompensa! 250.000 monedas de oro extra y su equipamiento", FontTypeNames.FONTTYPE_INFO
60          'WriteShortMsj UserIndex, 55, FontTypeNames.FONTTYPE_INFO, , , , 250000
70          .Stats.Gld = .Stats.Gld + 250000
                  
80      End If
              
90      MiObj.ObjIndex = 899
100         MiObj.Amount = 1
                              
110         If Not MeterItemEnInventario(UserIndex, MiObj) Then
120             Call TirarItemAlPiso(.Pos, MiObj)
130         End If
                              
140         MiObj.ObjIndex = 900
150         MiObj.Amount = 1
                              
160         If Not MeterItemEnInventario(UserIndex, MiObj) Then
170             Call TirarItemAlPiso(.Pos, MiObj)
180         End If
              
190         WriteUpdateGold UserIndex
              
200         .Stats.TorneosGanados = .Stats.TorneosGanados + 1
210     End With
          
220     Exit Sub

error:
230     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CastleMode_Premio()"
End Sub

' FIN EVENTO CASTLE MODE #####################################

' ###################### EVENTO DAGA RUSA ###########################
Public Sub InitDagaRusa(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim LoopC    As Integer

    Dim NpcIndex As Integer

    Dim Pos      As WorldPos
          
    Dim Num      As Integer
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
                Call WriteUserInEvent(.Users(LoopC).ID)
50              Call WarpUserChar(.Users(LoopC).ID, 60, 21 + Num, 60, False)
60              Num = Num + 1
80          End If

90      Next LoopC
              
            Pos.Map = 60
            Pos.X = 21
            Pos.Y = 59
            
            NpcIndex = CrearNPC(704, Pos.Map, Pos)
          
140         If NpcIndex <> 0 Then
150             Npclist(NpcIndex).Movement = NpcDagaRusa
160             Npclist(NpcIndex).flags.SlotEvent = SlotEvent
170             Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed
180             .NpcIndex = NpcIndex
                  
190             DagaRusa_MoveNpc NpcIndex, True
200         End If
              
210         .TimeCount = 4
220     End With

230     Exit Sub

error:
240     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitDagaRusa()"
End Sub

Public Function DagaRusa_NextUser(ByVal SlotEvent As Byte) As Byte

10  On Error GoTo error

    Dim LoopC As Integer
          
20  DagaRusa_NextUser = 0
          
30  With Events(SlotEvent)

40      For LoopC = LBound(.Users()) To UBound(.Users())

50          If (.Users(LoopC).ID > 0) And (.Users(LoopC).Value = 0) Then
60              DagaRusa_NextUser = .Users(LoopC).ID

                '.Users(LoopC).Value = 1
70              Exit For

80          End If

90      Next LoopC

100     End With
              
110     Exit Function

error:
120     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : DagaRusa_NextUser()"
End Function

Public Sub DagaRusa_ResetRonda(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
          
10  With Events(SlotEvent)

20      For LoopC = LBound(.Users()) To UBound(.Users())
30          .Users(LoopC).Value = 0
40      Next LoopC
          
50  End With

End Sub

Private Sub DagaRusa_CheckWin(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim UserIndex As Integer

    Dim MiObj     As Obj

    Dim strReward As String
          
20  With Events(SlotEvent)

30      If .Inscribed = 1 Then
40          UserIndex = SearchLastUserEvent(SlotEvent)
50          Call PrizeUser(UserIndex, strReward)
                  
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DagaRusa» El ganador es " & UserList(UserIndex).Name & ". " & strReward, FontTypeNames.FONTTYPE_GUILD)

60          Call QuitarNPC(.NpcIndex)
70          CloseEvent SlotEvent
                  
80      End If

90  End With
          
100     Exit Sub

error:
110     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : DagaRusa_CheckWin()"
End Sub

Public Sub DagaRusa_AttackUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

10  On Error GoTo error

    Dim N    As Integer

    Dim Slot As Byte
          
20  With UserList(UserIndex)
              
30      N = 30
              
        Randomize

40      If RandomNumber(1, 100) <= N Then
              
            ' Sound
50          SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
            ' Fx
60          SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
            ' Cambio de Heading
70          ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, SOUTH
            'Apuñalada en el piso
80          'SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1000, DAMAGE_PUÑAL)
                  
90          WriteConsoleMsg UserIndex, "¡Has sido apuñalado!", FontTypeNames.FONTTYPE_FIGHT
                  
100             Slot = .flags.SlotEvent
                  
110             Call UserDie(UserIndex)
120             EventosDS.AbandonateEvent (UserIndex)
130             Call DagaRusa_CheckWin(Slot)
                  
140         Else
                ' Sound
150             SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
                ' Fx
160             SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
                ' Cambio de Heading
170             ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, SOUTH

180             WriteConsoleMsg UserIndex, "¡JA JA JA! La proxima te vas", FontTypeNames.FONTTYPE_FIGHT
                ' SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1000, DAMAGE_PUÑAL)
190         End If
              
200     End With

210     Exit Sub

error:
220     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : DagaRusa_AttackUser()"
End Sub

' FIN EVENTO DAGA RUSA ###########################################
Private Function SelectModalityDeathMatch(ByVal SlotEvent As Byte) As Integer

    Dim Random As Integer
          
10  Randomize
20  Random = RandomNumber(1, 8)
          
30  With Events(SlotEvent)

40      Select Case Random

            Case 1 ' Zombie
50              .CharBody = 11

60          Case 2 ' Golem
70              .CharBody = 11

80          Case 3 ' Araña
90              .CharBody = 42

100             Case 4 ' Asesino
110                 .CharBody = 11 '48

120             Case 5 'Medusa suprema
130                 .CharBody = 151

140             Case 6 'Dragón azul
150                 .CharBody = 42 '247

160             Case 7 'Viuda negra 185
170                 .CharBody = 185

180             Case 8 'Tigre salvaje
190                 .CharBody = 147
200         End Select

210     End With

End Function

' DEATHMATCH ####################################################
Private Sub InitDeathMatch(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim LoopC As Integer

    Dim Pos   As WorldPos
          
20  Call SelectModalityDeathMatch(SlotEvent)
          
30  With Events(SlotEvent)

40      For LoopC = LBound(.Users()) To UBound(.Users())

50          If .Users(LoopC).ID > 0 Then
60              .Users(LoopC).Team = LoopC
70              .Users(LoopC).Selected = 1
                      
80              ChangeBodyEvent SlotEvent, .Users(LoopC).ID, True
90              UserList(.Users(LoopC).ID).ShowName = False
                    UserList(.Users(LoopC).ID).flags.Mimetizado = 1
100                 RefreshCharStatus .Users(LoopC).ID
                      
                    Pos.Map = 60
                    Pos.X = RandomNumber(58, 84)
                    Pos.Y = RandomNumber(28, 44)
                    Call EventWarpUser(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y)
160             End If
              
170         Next LoopC
          
180         .TimeCount = 20
190     End With
          
200     Exit Sub

error:
210     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitDeathMatch()"
End Sub

Public Sub DeathMatch_UserDie(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

10  On Error GoTo error
            
    Dim strReward As String
            
20  AbandonateEvent (UserIndex)
              
30  If Events(SlotEvent).Inscribed = 1 Then
40      UserIndex = SearchLastUserEvent(SlotEvent)
50      Call PrizeUser(UserIndex, strReward)
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DeathMatch» El ganador es " & UserList(UserIndex).Name & ". " & strReward, FontTypeNames.FONTTYPE_GUILD)
60      CloseEvent SlotEvent
70  End If
          
80  Exit Sub

error:
90  LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : DeathMatch_UserDie()"
End Sub

' FIN DEATHMATCH ################################################
' EVENTOS DONDE LOS USUARIOS SE TRANSFORMAN EN CRIATURAS
Private Sub InitEventTransformation(ByVal SlotEvent As Byte, _
                                    ByVal NewBody As Integer, _
                                    ByVal NewHp As Integer, _
                                    ByVal Map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)
          
10  On Error GoTo error
          
    Dim LoopC        As Integer

    Dim UserSelected As Integer

    Dim Pos          As WorldPos
          
    Const Rango      As Byte = 4
          
20  With Events(SlotEvent)
30      .CharBody = NewBody
40      .CharHp = NewHp
              
50      For LoopC = LBound(.Users()) To UBound(.Users())

60          If .Users(LoopC).ID > 0 Then
70              .Users(LoopC).Team = 2
                      
80              Pos.Map = Map
90              Pos.X = RandomNumber(X - Rango, X + Rango)
100                 Pos.Y = RandomNumber(Y - Rango, Y + Rango)
                  
110                 Call ClosestLegalPos(Pos, Pos)
120                 Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, True)
                      
130             End If

140         Next LoopC
              
150         Transformation_SelectionUser SlotEvent
160     End With
          
170     Exit Sub

error:
180     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitEventTransformation()"
End Sub

Private Function Transformation_SelectionUser(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim LoopC As Integer

20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())
40          Transformation_SelectionUser = RandomNumber(LBound(.Users()), UBound(.Users()))
                  
50          If .Users(Transformation_SelectionUser).ID > 0 And .Users(Transformation_SelectionUser).Selected = 0 Then

60              Exit For

70          End If

80      Next LoopC
              
90      .Users(Transformation_SelectionUser).Selected = 1
100         .Users(Transformation_SelectionUser).Team = 1
                          
110         Call ChangeHpEvent(.Users(Transformation_SelectionUser).ID)
120         Call ChangeBodyEvent(SlotEvent, .Users(Transformation_SelectionUser).ID, IIf(.Modality = Minotauro, False, True))
130     End With

140     Exit Function

error:
150     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Transformation_SelectionUser()"
End Function

Public Sub Transformation_UserDie(ByVal UserIndex As Integer, _
                                  ByVal AttackerIndex As Integer)

10  On Error GoTo error

    Dim SlotEvent As Byte

    Dim Exituser  As Boolean
          
20  With UserList(UserIndex)
30      SlotEvent = .flags.SlotEvent
40      AbandonateEvent UserIndex
              
50      Transformation_CheckWin UserIndex, SlotEvent, AttackerIndex
60  End With

70  Exit Sub

error:
80  LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Transformation_UserDie()"
End Sub

Private Function Transformation_SearchUserSelected(ByVal SlotEvent As Byte) As Integer

10  On Error GoTo error

    Dim LoopC As Integer
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
50              If .Users(LoopC).Selected = 1 Then
60                  Transformation_SearchUserSelected = LoopC
70              End If
80          End If

90      Next LoopC

100     End With
          
110     Exit Function

error:
120     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Transformation_SearchUserSelected()"
End Function

Public Sub Transformation_CheckWin(ByVal UserIndex As Integer, _
                                   ByVal SlotEvent As Byte, _
                                   Optional ByVal AttackerIndex As Integer = 0)

10  On Error GoTo error
    
    ' VER LAUTARO
    Dim IsSelected As Boolean

    Dim tUser      As Integer

20

30  With Events(SlotEvent)

40      If .Inscribed = 1 Then
50          tUser = SearchLastUserEvent(SlotEvent)
60

70          If .Users(UserList(tUser).flags.SlotUserEvent).Selected = 1 Then IsSelected = True
                
80          Transformation_Premio tUser, IsSelected, 250000
90
100             CloseEvent SlotEvent

110             Exit Sub

120         End If

130
        
            If AttackerIndex <> 0 Then

                'Significa que hay más de un usuario. Por lo tanto podría haber muerto el bicho transformado
140             If UserList(UserIndex).flags.SlotUserEvent = Transformation_SearchUserSelected(SlotEvent) Then
150                 Transformation_Premio AttackerIndex, False, 250000
160
170                 CloseEvent SlotEvent
180             End If
            End If

190     End With
    
200     Exit Sub

error:
210     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Transformation_CheckWin() at line " & Erl
End Sub

Private Sub Transformation_Premio(ByVal UserIndex As Integer, _
                                  ByVal IsSelected As Boolean, _
                                  ByVal Gld As Long)
                                    
10  On Error GoTo error

20

    Dim UserWin As Integer
    
30  With UserList(UserIndex)

        Dim SlotEvent As Byte

40      SlotEvent = .flags.SlotEvent
        
50      If IsSelected Then
60          .Stats.Gld = .Stats.Gld + (Gld * 2)
            WriteConsoleMsg UserIndex, "Has recibido " & (Gld * 2) & " por haber aniquilado a todos los usuarios.", FontTypeNames.FONTTYPE_INFO
70          SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, Events(SlotEvent).Modality) & "» Ha logrado derrotar a todos los participantes. Felicitaciones para " & .Name & " quien fue escogido como " & strModality(SlotEvent, Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_GUILD)
80          'WriteShortMsj UserIndex, 58, FontTypeNames.FONTTYPE_INFO, , , , (Gld * 2)

90      Else
100             .Stats.Gld = .Stats.Gld + Gld
                WriteConsoleMsg UserIndex, "Has recibido " & Gld & " por haber aniquilado a " & strModality(SlotEvent, Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_INFO
110             'WriteShortMsj UserIndex, 59, FontTypeNames.FONTTYPE_INFO, , , , Gld, strModality(SlotEvent, Events(SlotEvent).Modality)
120             SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, Events(SlotEvent).Modality) & "» Felicitaciones para " & .Name & " quien derrotó a " & strModality(SlotEvent, Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_GUILD)

130         End If
        
140         WriteUpdateGold UserIndex
        
150         .Stats.TorneosGanados = .Stats.TorneosGanados + 1
    
160     End With
    
170     Exit Sub

error:
180     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Transformation_Premio() AT LINE: " & Erl
End Sub

' FIN EVENTOS DONDE LOS USUARIOS SE TRANSFORMAN EN CRIATURAS

' ARACNUS #######################################################

Public Sub Aracnus_Veneno(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

10  On Error GoTo error

    ' El personaje transformado en Aracnus, tiene 10% de probabilidad de envenenar a la víctima y dejarla fuera del torneo.

    Const N As Byte = 10
          
20  With UserList(AttackerIndex)

30      If RandomNumber(1, 100) <= 10 Then
            WriteConsoleMsg VictimIndex, "Has sido envenenado por Aracnus, has muerto de inmediato por su veneno letal.", FontTypeNames.FONTTYPE_FIGHT
40          'WriteShortMsj VictimIndex, 60, FontTypeNames.FONTTYPE_FIGHT
50          Call UserDie(VictimIndex)
                  
60          Transformation_CheckWin VictimIndex, .flags.SlotEvent, AttackerIndex
70      End If
          
80  End With
          
90  Exit Sub

error:
100     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Aracnus_Veneno()"
End Sub

Public Sub Minotauro_Veneno(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

10  On Error GoTo error

    ' El personaje transformado en Minotauro, tiene 10% de posibilidad de dar un golpe mortal
    Const N As Byte = 10
          
20  With UserList(AttackerIndex)

30      If RandomNumber(1, 100) <= 10 Then
            WriteConsoleMsg VictimIndex, "¡El minotauro ha logrado paralizar tu cuerpo con su dosis de veneno. Has quedado afuera del evento.", FontTypeNames.FONTTYPE_FIGHT
40          'WriteShortMsj VictimIndex, 61, FontTypeNames.FONTTYPE_FIGHT
50          Call UserDie(VictimIndex)
                  
60          Transformation_CheckWin VictimIndex, .flags.SlotEvent, AttackerIndex
              
70      End If
          
80  End With
          
90  Exit Sub

error:
100     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Minotauro_Veneno()"
End Sub

' FIN ARACNUS ###################################################

' EVENTO BUSQUEDA '
Private Sub InitBusqueda(ByVal SlotEvent As Byte)

10  On Error GoTo error
          
    Dim LoopC As Integer

    Dim Pos   As WorldPos

20  With Events(SlotEvent)


30      For LoopC = 1 To 20
40          Event_MakeObj 1037, 62, RandomNumber(20, 80), RandomNumber(20, 80), 1
50      Next LoopC
              
60      For LoopC = LBound(.Users()) To UBound(.Users())

70          If .Users(LoopC).ID > 0 Then
80              Pos.Map = 62
90              Pos.X = RandomNumber(50, 60)
100                 Pos.Y = RandomNumber(50, 60)
                      
110                 Call ClosestLegalPos(Pos, Pos)
120                 Call WarpUserChar(.Users(LoopC).ID, Pos.Map, Pos.X, Pos.Y, True)
130             End If

140         Next LoopC
              
150         .TimeFinish = 60
          
160     End With
          
170     Exit Sub

error:
180     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitBusqueda()"
End Sub

Private Sub Event_MakeObj(ByVal ObjIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal ObjEvent As Byte)

10  On Error GoTo error

    ' Creamos un objeto en el mapa de búsqueda.
          
    Dim Pos As WorldPos

    Dim Obj As Obj
          
20  Pos.Map = Map
30  Pos.X = X
40  Pos.Y = Y
50  ClosestStablePos Pos, Pos
          
60  Obj.ObjIndex = ObjIndex
70  Obj.Amount = 1
80  Call MakeObj(Obj, Pos.Map, Pos.X, Pos.Y)
90  MapData(Pos.Map, Pos.X, Pos.Y).ObjEvent = ObjEvent
          
100     Exit Sub

error:
110     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Event_MakeObj()"
End Sub

Private Sub Busqueda_SearchWin(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim UserIndex   As Integer

    Dim CopyUsers() As tUserEvent

    Dim strReward   As String
          
    Dim LoopX As Integer
    Dim LoopY As Integer
    
20  With Events(SlotEvent)
        For LoopX = XMinMapSize To XMaxMapSize
            For LoopY = YMinMapSize To YMaxMapSize
    
                If InMapBounds(62, LoopX, LoopY) Then
                    
                    If MapData(62, LoopX, LoopY).ObjEvent = 1 Then
                        MapData(62, LoopX, LoopY).ObjEvent = 0
                        EraseObj 10000, 62, LoopX, LoopY
                    End If
                    
                End If
    
            Next LoopY
        Next LoopX
        
30      Event_OrdenateUsersValue SlotEvent, CopyUsers
              
40      UserIndex = CopyUsers(1).ID
              
50      If UserIndex > 0 Then
                  
            Call PrizeUser(UserIndex, strReward)
                  
90          SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Busqueda de objetos» El ganador de la búsqueda de objetos es " & UserList(UserIndex).Name & ". Felicitaciones! " & strReward & vbCrLf & "Tabla final de posiciones: " & vbCrLf & Event_GenerateTablaPos(SlotEvent, CopyUsers), FontTypeNames.FONTTYPE_GUILD)
              
110         CloseEvent SlotEvent

        End If
    End With
130     Exit Sub

error:
140     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Busqueda_SearchWin()"
End Sub

Private Function Busqueda_UserRecolectedObj(ByVal SlotEvent As Byte) As Integer

10  On Error GoTo error

    Dim LoopC As Integer
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())
                  
40          If .Users(LoopC).ID > 0 Then
50              If Busqueda_UserRecolectedObj = 0 Then Busqueda_UserRecolectedObj = LoopC
60              If .Users(LoopC).Value > .Users(Busqueda_UserRecolectedObj).Value Then
70                  Busqueda_UserRecolectedObj = LoopC
80              End If
90          End If
                      
100         Next LoopC
              
110         Busqueda_UserRecolectedObj = .Users(Busqueda_UserRecolectedObj).ID
120     End With
          
130     Exit Function

error:
140     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Busqueda_UserRecolectedObj()"
End Function

Public Sub Busqueda_GetObj(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte)

10  On Error GoTo error

20  With Events(SlotEvent)
30      .Users(SlotUserEvent).Value = .Users(SlotUserEvent).Value + 1
              
        WriteConsoleMsg .Users(SlotUserEvent).ID, "Has recolectado un objeto del piso. En total llevas " & .Users(SlotUserEvent).Value & " objetos recolectados. Sigue así!", FontTypeNames.FONTTYPE_INFO
40      'WriteShortMsj .Users(SlotUserEvent).Id, 63, FontTypeNames.FONTTYPE_INFO, .Users(SlotUserEvent).Value
50      Event_MakeObj 1037, 62, RandomNumber(30, 80), RandomNumber(30, 80), 1
60  End With

70  Exit Sub

error:
80  LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Busqueda_GetObj()"
End Sub

' ENFRENTAMIENTOS ###############################################

Private Sub InitFights(ByVal SlotEvent As Byte)

10  On Error GoTo error
          
20  With Events(SlotEvent)
30      Fight_SelectedTeam SlotEvent
40      Fight_Combate SlotEvent
50  End With

60  Exit Sub

error:
70  LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitFights()"
End Sub

Private Sub Fight_SelectedTeam(ByVal SlotEvent As Byte)
          
10  On Error GoTo error

    ' En los enfrentamientos utilizamos este procedimiento para seleccionar los grupos o bien el usuario queda solo por 1vs1.
    Dim LoopX   As Integer

    Dim LoopY   As Integer

    Dim Team    As Byte

    Dim TeamSTR As String
          
20  Team = 1
          
30  With Events(SlotEvent)

40      For LoopX = LBound(.Users()) To UBound(.Users()) Step .TeamCant
50          For LoopY = 0 To (.TeamCant - 1)
60              .Users(LoopX + LoopY).Team = Team
70          Next LoopY
                  
80          Team = Team + 1
90      Next LoopX

100     End With
          
110     Exit Sub

error:
120     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_SelectedTeam()"
End Sub

Private Sub Fight_WarpTeam(ByVal SlotEvent As Byte, _
                           ByVal ArenaSlot As Byte, _
                           ByVal TeamEvent As Byte, _
                           ByVal IsContrincante As Boolean, _
                           ByRef StrTeam As String)

10  On Error GoTo error

    Dim LoopC   As Integer

    Dim strTemp As String, strTemp1 As String, strTemp2 As String
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 And .Users(LoopC).Team = TeamEvent Then
                LogEventos "Usuario: " & UserList(.Users(LoopC).ID).Name & ". Team: " & TeamEvent
                
50              If IsContrincante Then
60                  Call EventWarpUser(.Users(LoopC).ID, MapEvent.Fight(ArenaSlot).Map, MapEvent.Fight(ArenaSlot).X + MAP_TILE_VS, MapEvent.Fight(ArenaSlot).Y + MAP_TILE_VS)
                          
                    ' / Update color char team
70                  UserList(.Users(LoopC).ID).flags.FightTeam = 2
                          
80                  RefreshCharStatus (.Users(LoopC).ID)
90              Else
100                     Call EventWarpUser(.Users(LoopC).ID, MapEvent.Fight(ArenaSlot).Map, MapEvent.Fight(ArenaSlot).X, MapEvent.Fight(ArenaSlot).Y)
                          
                        ' / Update color char team
110                     UserList(.Users(LoopC).ID).flags.FightTeam = 1
120                     RefreshCharStatus (.Users(LoopC).ID)
130                 End If
                      
140                 If StrTeam = vbNullString Then
150                     StrTeam = UserList(.Users(LoopC).ID).Name
160                 Else
170                     StrTeam = StrTeam & "-" & UserList(.Users(LoopC).ID).Name
180                 End If
                      
190                 .Users(LoopC).Value = 1
200                 .Users(LoopC).MapFight = ArenaSlot
                      
210                 UserList(.Users(LoopC).ID).Counters.TimeFight = 10
220                 Call WriteUserInEvent(.Users(LoopC).ID)
230             End If

240         Next LoopC

250     End With
          
260     Exit Sub

error:
270     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_WarpTeam()"
End Sub

Private Function Fight_Search_Enfrentamiento(ByVal UserIndex As Integer, _
                                             ByVal UserTeam As Byte, _
                                             ByVal SlotEvent As Byte) As Byte

10  On Error GoTo error

    ' Chequeamos que tengamos contrincante para luchar.
    Dim LoopC As Integer
          
20  Fight_Search_Enfrentamiento = 0
          
30  With Events(SlotEvent)

40      For LoopC = LBound(.Users()) To UBound(.Users())

50          If .Users(LoopC).ID > 0 And .Users(LoopC).Value = 0 Then
60              If .Users(LoopC).ID <> UserIndex And .Users(LoopC).Team <> UserTeam Then
70                  Fight_Search_Enfrentamiento = .Users(LoopC).Team

80                  Exit For

90              End If
100             End If

110         Next LoopC
          
120     End With
          
130     Exit Function

error:
140     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_Search_Enfrentamiento()"
End Function

Private Sub NewRound(ByVal SlotEvent As Byte)

    Dim LoopC As Long

    Dim Count As Long
          
10  With Events(SlotEvent)
20      Count = 0
              
30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then

                ' Hay esperando
50              If .Users(LoopC).Value = 0 Then

60                  Exit Sub

70              End If
                      
                ' Hay luchando
80              If .Users(LoopC).MapFight > 0 Then

90                  Exit Sub

100                 End If
110             End If

120         Next LoopC
              
130         For LoopC = LBound(.Users()) To UBound(.Users())
140             .Users(LoopC).Value = 0
150         Next LoopC

            LogEventos "Se reinicio la informacion de los fights()"
              
160     End With

End Sub

Private Sub Fight_Combate(ByVal SlotEvent As Byte)

10  On Error GoTo error

    ' Buscamos una arena disponible y mandamos la mayor cantidad de usuarios disponibles.
    Dim LoopC       As Integer

    Dim FreeArena   As Byte

    Dim OponentTeam As Byte

    Dim strTemp     As String

    Dim strTeam1    As String

    Dim strTeam2    As String
          
20  With Events(SlotEvent)
cheking:

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 And .Users(LoopC).Value = 0 Then
50              FreeArena = FreeSlotArena()
                      
60              If FreeArena > 0 Then
70                  OponentTeam = Fight_Search_Enfrentamiento(.Users(LoopC).ID, .Users(LoopC).Team, SlotEvent)
                          
80                  If OponentTeam > 0 Then
90                      StatsEvent .Users(LoopC).ID
100                         Fight_WarpTeam SlotEvent, FreeArena, .Users(LoopC).Team, False, strTeam1
110                         Fight_WarpTeam SlotEvent, FreeArena, OponentTeam, True, strTeam2
120                         MapEvent.Fight(FreeArena).Run = True
                              
130                         strTemp = "Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» "
140                         strTemp = strTemp & strTeam1 & " vs " & strTeam2
150                         SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_GUILD)
                              
160                         strTemp = vbNullString
170                         strTeam1 = vbNullString
180                         strTeam2 = vbNullString
                              
190                     Else
                            ' Pasa de ronda automaticamente
200                         .Users(LoopC).Value = 1
210                         WriteConsoleMsg .Users(LoopC).ID, "Hemos notado que no tienes un adversario. Pasaste a la siguiente ronda.", FontTypeNames.FONTTYPE_INFO
220                         NewRound SlotEvent
                            GoTo cheking:
230                     End If
240                 End If
250             End If

260         Next LoopC
              
270     End With
          
280     Exit Sub

error:
290     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_Combate()"
End Sub

Private Sub ResetValue(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
          
10  With Events(SlotEvent)

20      For LoopC = LBound(.Users()) To UBound(.Users())
30          .Users(LoopC).Value = 0
40      Next LoopC

50  End With

End Sub

Private Function CheckTeam_UserDie(ByVal SlotEvent As Integer, _
                                   ByVal TeamUser As Byte) As Boolean

10  On Error GoTo error

    Dim LoopC As Integer

    ' Encontramos a uno del Team vivo, significa que no hay terminación del duelo.
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
50              If .Users(LoopC).Team = TeamUser Then
60                  If UserList(.Users(LoopC).ID).flags.Muerto = 0 Then
70                      CheckTeam_UserDie = False

80                      Exit Function

90                  End If
100                 End If
110             End If

120         Next LoopC
              
130         CheckTeam_UserDie = True
          
140     End With
          
150     Exit Function

error:
160     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CheckTeam_UserDie()"
End Function

Private Sub Team_UserDie(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte)

10  On Error GoTo error

    Dim LoopC As Integer

20  With Events(SlotEvent)
              
30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
50              If .Users(LoopC).Team = TeamSlot Then
60                  AbandonateEvent .Users(LoopC).ID
70              End If
80          End If

90      Next LoopC
          
100     End With
          
110     Exit Sub

error:
120     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Team_UserDie()"
End Sub

Public Function Fight_CheckContinue(ByVal UserIndex As Integer, _
                                    ByVal SlotEvent As Byte, _
                                    ByVal TeamSlot As Byte) As Boolean
    ' Esta función devuelve un TRUE cuando el enfrentamiento puede seguir.
          
    Dim LoopC As Integer, cant As Integer
          
10  With Events(SlotEvent)
              
20      Fight_CheckContinue = False
              
30      For LoopC = LBound(.Users()) To UBound(.Users())

            ' User válido
40          If .Users(LoopC).ID > 0 And .Users(LoopC).ID <> UserIndex Then
50              If .Users(LoopC).Team = TeamSlot Then
60                  If UserList(.Users(LoopC).ID).flags.Muerto = 0 Then
70                      Fight_CheckContinue = True

80                      Exit For

90                  End If
100                 End If
110             End If

120         Next LoopC

130     End With
          
140     Exit Function

error:
150     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Team_CheckContinue()"
End Function

Public Sub Fight_WinForzado(ByVal UserIndex As Integer, _
                            ByVal SlotEvent As Byte, _
                            ByVal MapFight As Byte)

10  On Error GoTo error
          
    Dim LoopC      As Integer

    Dim strTempWin As String

    Dim TeamWin    As Byte
          
20  With Events(SlotEvent)

        LogEventos "El personaje " & UserList(UserIndex).Name & " deslogeó en lucha."
              
30      For LoopC = LBound(.Users()) To UBound(.Users())

40          With .Users(LoopC)

50              If .ID > 0 And UserIndex <> .ID Then
60                  If .MapFight = MapFight Then
70                      If strTempWin = vbNullString Then
80                          strTempWin = UserList(.ID).Name
90                      Else
100                             strTempWin = strTempWin & "-" & UserList(.ID).Name
110                         End If
                              
                            '.value = 0
130                         .MapFight = 0
                              
140                         EventWarpUser .ID, 60, 30, 21
                            WriteConsoleMsg .ID, "Felicitaciones. Has ganado el enfrentamiento", FontTypeNames.FONTTYPE_INFO
                            LogEventos "El personaje " & UserList(.ID).Name & " ha ganado el enfrentamiento"
                              
150                         'WriteShortMsj .Id, 64, FontTypeNames.FONTTYPE_INFO

                            ' / Update color char team
160                         UserList(.ID).flags.FightTeam = 0
170                         RefreshCharStatus (.ID)
180                         TeamWin = .Team
190                     End If
200                 End If

210             End With

220         Next LoopC

            MapEvent.Fight(MapFight).Run = False
              
230         If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Duelo ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
              
            ' Nos fijamos si resetea el Value
240         Call NewRound(SlotEvent)
              
            ' Nos fijamos si eran los últimos o si podemos mandar otro combate..
250         If TeamCant(SlotEvent, TeamWin) = .Inscribed Then
260             Fight_SearchTeamWin SlotEvent, TeamWin
270             CloseEvent SlotEvent
280         Else
290             Fight_Combate SlotEvent
300         End If
          
310     End With
          
320     Exit Sub

error:
330     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_WinForzado()"
End Sub

Private Sub StatsEvent(ByVal UserIndex As Integer)

10  On Error GoTo error

20  With UserList(UserIndex)

30      If .flags.Muerto Then
40          Call RevivirUsuario(UserIndex)

50          Exit Sub

60      End If
              
70      .Stats.MinHp = .Stats.MaxHp
80      .Stats.MinMan = .Stats.MaxMan
90      .Stats.MinAGU = 100
100         .Stats.MinHam = 100
              
110         WriteUpdateUserStats UserIndex
          
120     End With
          
130     Exit Sub

error:
140     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StatsEvent()"
End Sub

Private Function SearchTeamAttacker(ByVal TeamUser As Byte)

End Function

Public Sub Fight_UserDie(ByVal SlotEvent As Byte, _
                         ByVal SlotUserEvent As Byte, _
                         ByVal AttackerIndex As Integer)

10  On Error GoTo error

    Dim TeamSlot   As Byte

    Dim LoopC      As Integer

    Dim strTempWin As String

    Dim TeamWin    As Byte

    Dim MapFight   As Byte
    
    ' Aca se hace que el que gané no siga luchando sino que espere.
    
20  With Events(SlotEvent)
30      TeamSlot = .Users(SlotUserEvent).Team
40      TeamWin = .Users(UserList(AttackerIndex).flags.SlotUserEvent).Team
        
50      If CheckTeam_UserDie(SlotEvent, TeamSlot) = False Then Exit Sub
        
60      For LoopC = LBound(.Users()) To UBound(.Users())

70          If .Users(LoopC).ID > 0 Then

80              With .Users(LoopC)

90                  If .Team = TeamWin Then
100                         StatsEvent .ID
110

120                         If strTempWin = vbNullString Then
130                             strTempWin = UserList(.ID).Name
140                         Else
150                             strTempWin = strTempWin & "-" & UserList(.ID).Name
160                         End If
                            
                            MapFight = .MapFight
170
                            
                            '.value = 0
180                         .MapFight = 0
190                         EventWarpUser .ID, 60, 30, 21
                            WriteConsoleMsg .ID, "Felicitaciones. Has ganado el duelo", FontTypeNames.FONTTYPE_INFO
200                         'WriteShortMsj .Id, 64, FontTypeNames.FONTTYPE_INFO
                           
                            ' / Update color char team
210                         UserList(.ID).flags.FightTeam = 0
220                         RefreshCharStatus (.ID)

                            If UserList(.ID).flags.Muerto Then RevivirUsuario (.ID)
230                     End If

240                 End With

250             End If

260         Next LoopC
        
            MapEvent.Fight(MapFight).Run = False
        
            ' Abandono del user/team
270         Team_UserDie SlotEvent, TeamSlot
        
280         If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Duelo ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
        
            ' // Se fija de poder pasar a la siguiente ronda o esperar a los combates que faltan.
290         Call NewRound(SlotEvent)
        
            ' Si la cantidad es igual al inscripto quedó final.
300         If TeamCant(SlotEvent, TeamWin) = .Inscribed Then
310             Fight_SearchTeamWin SlotEvent, TeamWin
320             CloseEvent SlotEvent
330         Else
340             Fight_Combate SlotEvent
350         End If
        
360     End With
    
370     Exit Sub

error:
380     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_UserDie()" & " AT LINE: " & Erl
End Sub

Private Function TeamCant(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte) As Byte

10  On Error GoTo error

    ' Devuelve la cantidad de miembros que tiene un clan
    Dim LoopC As Integer
          
20  TeamCant = 0
          
30  With Events(SlotEvent)

40      For LoopC = LBound(.Users()) To UBound(.Users())

50          If .Users(LoopC).Team = TeamSlot Then
60              TeamCant = TeamCant + 1
70          End If

80      Next LoopC

90  End With
          
100     Exit Function

error:
110     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : TeamCant()"
End Function

Private Sub Fight_SearchTeamWin(ByVal SlotEvent As Byte, ByVal TeamWin As Byte)

10  On Error GoTo error

    Dim LoopC     As Integer

    Dim strTemp   As String

    Dim strReward As String
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 And .Users(LoopC).Team = TeamWin Then
                WriteConsoleMsg .Users(LoopC).ID, "Has ganado el evento. ¡Felicitaciones!", FontTypeNames.FONTTYPE_INFO
50              'WriteShortMsj .Users(LoopC).Id, 65, FontTypeNames.FONTTYPE_INFO
                      
60              PrizeUser .Users(LoopC).ID, strReward
                      
70              If strTemp = vbNullString Then
80                  strTemp = UserList(.Users(LoopC).ID).Name
90              Else
100                     strTemp = strTemp & ", " & UserList(.Users(LoopC).ID).Name
110                 End If
120             End If

130         Next LoopC
          
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & .TeamCant & "vs" & .TeamCant & "» Evento terminado. Felicitamos a " & strTemp & " por haber ganado el torneo." & vbCrLf & strReward, FontTypeNames.FONTTYPE_INFOBOLD)
          
250     End With
          
260     Exit Sub

error:
270     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Fight_SearchTeamWin()"
End Sub

' ############################## USUARIO UNSTOPPABLE ###########################################
Public Sub InitUnstoppable(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim LoopC As Integer
          
20  With Events(SlotEvent)

30      For LoopC = LBound(.Users()) To UBound(.Users())

40          If .Users(LoopC).ID > 0 Then
50              EventWarpUser .Users(LoopC).ID, 64, RandomNumber(30, 54), RandomNumber(25, 39)
                      
60          End If

70      Next LoopC
              
80      .TimeCount = 10
90      .TimeFinish = 60 + .TimeCount
100     End With
          
110     Exit Sub

error:
120     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitUnstoppable()"
End Sub

Private Function Event_GenerateTablaPos(ByVal SlotEvent As Byte, _
                                        ByRef CopyUsers() As tUserEvent) As String

    Dim LoopC As Integer
          
10  With Events(SlotEvent)

20      For LoopC = LBound(.Users()) To UBound(.Users())

30          If CopyUsers(LoopC).ID > 0 Then
40              Event_GenerateTablaPos = Event_GenerateTablaPos & LoopC & "° »» " & UserList(CopyUsers(LoopC).ID).Name & " (" & CopyUsers(LoopC).Value & ")" & vbCrLf
50          End If

60      Next LoopC

70  End With
          
End Function

Private Sub Unstoppable_UserWin(ByVal SlotEvent As Byte)

10  On Error GoTo error

    Dim UserIndex   As Integer

    Dim strTemp     As String
    
    Dim strReward As String

    Dim CopyUsers() As tUserEvent
          
20  Event_OrdenateUsersValue SlotEvent, CopyUsers
          
30  UserIndex = CopyUsers(1).ID
          
40  With UserList(UserIndex)
50      'WriteShortMsj UserIndex, 68, FontTypeNames.FONTTYPE_GUILD, Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Value
       ' WriteConsoleMsg UserIndex, "Felicitaciones. Tus " & Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Value & " asesinatos han hecho que ganes el evento. Aquí tienes 500.000 monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO
60      '.Stats.Gld = .Stats.Gld + 350000
70      '.Stats.TorneosGanados = .Stats.TorneosGanados + 1
80     ' WriteUpdateGold UserIndex
        
        PrizeUser UserIndex, strReward
        
90      SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Usuario Unstoppable» El ganador del evento es " & .Name & " con " & Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Value & " asesinatos." & strReward & vbCrLf & "Tabla de posiciones: " & Event_GenerateTablaPos(SlotEvent, CopyUsers), FontTypeNames.FONTTYPE_GUILD)
                  
            'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(69, FontTypeNames.FONTTYPE_GUILD, Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).value, , , , Event_GenerateTablaPos)
100         CloseEvent SlotEvent
110     End With
          
120     Exit Sub

error:
130     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Unstoppable_UserWin()"
End Sub

Private Sub Event_OrdenateUsersValue(ByVal SlotEvent As Byte, _
                                     ByRef CopyUsers() As tUserEvent)

10  On Error GoTo error

    ' Utilizados para buscar ganador según VALUE
    Dim LoopX As Integer

    Dim LoopY As Integer

    Dim aux   As tUserEvent

    ' Dim CopyUsers() As tUserEvent
    
20  With Events(SlotEvent)
        ' Utilizamos la copia para no dañar archivos originales
30      ReDim CopyUsers(LBound(.Users()) To UBound(.Users())) As tUserEvent
        
40      For LoopY = LBound(.Users()) To UBound(.Users())
50          CopyUsers(LoopY) = .Users(LoopY)
60      Next LoopY
        
70      For LoopX = LBound(CopyUsers()) To UBound(CopyUsers())
80          For LoopY = LBound(CopyUsers()) To UBound(CopyUsers()) - LoopX

90              If CopyUsers(LoopY).ID > 0 Then
110                     If CopyUsers(LoopY).Value < CopyUsers(LoopY + 1).Value Then
                            
120                             aux = CopyUsers(LoopY)
                            
130                             CopyUsers(LoopY) = CopyUsers(LoopY + 1)
140                             CopyUsers(LoopY + 1) = aux
150                         End If
170                 End If

180             Next LoopY
190         Next LoopX

200     End With
    
210     Exit Sub

error:
220     LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : Event_OrdenateUsersValue()"
End Sub

' Eventos que cambien la clase del personaje
Public Sub Event_ChangeUser(ByVal UserIndex As Integer, ByVal SlotEvent As Byte)
    
    Dim TempMan As Integer

    Dim TempHit As Integer

    Dim TempHp  As Single

    Dim A       As Long
    
    Dim Clase As eClass
    Dim Raza As eRaza
    
    With UserList(UserIndex)
        Call Event_UserSaveOrig(UserIndex)
        Call ResetUserSpells(UserIndex)
        
        .Clase = IIf((Events(SlotEvent).ChangeClass > 0), Events(SlotEvent).ChangeClass, .Clase)
        .Raza = IIf((Events(SlotEvent).ChangeRaze > 0), Events(SlotEvent).ChangeRaze, .Raza)
        .Stats.Gld = 50000000
        .Stats.Eldhir = 10000
        
        .Stats.UserAtributos(eAtributos.Fuerza) = 18 + ModRaza(.Raza).Fuerza
        .Stats.UserAtributos(eAtributos.Agilidad) = 18 + ModRaza(.Raza).Agilidad
        .Stats.UserAtributos(eAtributos.Inteligencia) = 18 + ModRaza(.Raza).Inteligencia
        .Stats.UserAtributos(eAtributos.Carisma) = 18 + ModRaza(.Raza).Carisma
        .Stats.UserAtributos(eAtributos.Constitucion) = 18 + ModRaza(.Raza).Constitucion
    
        For A = 2 To STAT_MAXELV
            TempHp = TempHp + Balance_AumentoHP_Initial(.Clase, .Stats.UserAtributos(eAtributos.Constitucion))
            TempMan = TempMan + Balance_AumentoMANA(.Clase, .Raza, TempMan)
            TempHit = TempHit + Balance_AumentoHIT(UserIndex, A)
        Next A
        
        For A = 1 To NUMSKILLS
            .Stats.UserSkills(A) = 100
        Next A
        
        .Stats.MaxHp = 20 + TempHp
        .Stats.MaxMan = TempMan
        .Stats.MaxHit = 2 + TempHit

        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinHit = .Stats.MaxHit
    
        'If .Stats.MaxMan > 0 Then
            'Call Event_UserSpells(UserIndex)
        'End If
        
        Call UpdateUserHechizos(True, UserIndex, 0)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
    End With

End Sub

' Guardamos la información original
Public Sub Event_UserSaveOrig(ByVal UserIndex As Integer)

    Dim A As Long
    
    With UserList(UserIndex)
        .OldInfo.Clase = .Clase
        .OldInfo.Raza = .Raza
        .OldInfo.MaxHp = .Stats.MaxHp
        .OldInfo.MaxMan = .Stats.MaxMan
        .OldInfo.GldRed = .Stats.Gld
        .OldInfo.GldBlue = .Stats.Eldhir
        .OldInfo.MinHit = .Stats.MinHit
        .OldInfo.MaxHit = .Stats.MaxHit
        
        For A = 1 To NUMSKILLS
            .OldInfo.UserSkills(A) = .Stats.UserSkills(A)
        Next A
        
        For A = 1 To MAXUSERHECHIZOS
            .OldInfo.UserSpell(A) = .Stats.UserHechizos(A)
        Next A
        
        ' Chequeo si tomo poción de dopa
        If .flags.TomoPocion Then
            .flags.TomoPocion = 0
            .flags.TipoPocion = 0
            .flags.DuracionEfecto = 0
            
            For A = 1 To NUMATRIBUTOS
                .Stats.UserAtributos(A) = .Stats.UserAtributosBackUP(A)
            Next A

        End If
        
        For A = 1 To NUMATRIBUTOS
            .OldInfo.UserAtributos(A) = .Stats.UserAtributos(A)
        Next A
        
    End With

End Sub

' Volvemos al personaje a su situación inicial.
Public Sub Event_UserResetClass(ByVal UserIndex As Integer)

    Dim A As Long
   
    With UserList(UserIndex)
        .Clase = .OldInfo.Clase
        .Raza = .OldInfo.Raza
        .Stats.MaxHp = .OldInfo.MaxHp
        .Stats.MaxMan = .OldInfo.MaxMan
        .Stats.Gld = .OldInfo.GldRed
        .Stats.Eldhir = .OldInfo.GldBlue
        .Stats.MinHit = .OldInfo.MinHit
        .Stats.MaxHit = .OldInfo.MaxHit
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMan = .Stats.MaxMan
        
        .flags.DuracionEfecto = 0
        .flags.TomoPocion = False
        .flags.TipoPocion = 0
                
        For A = 1 To MAXUSERHECHIZOS
            .Stats.UserHechizos(A) = .OldInfo.UserSpell(A)
        Next A

        For A = 1 To NUMATRIBUTOS
            .Stats.UserAtributos(A) = .OldInfo.UserAtributos(A)
        Next A
        
        For A = 1 To NUMSKILLS
            .Stats.UserSkills(A) = .OldInfo.UserSkills(A)
        Next A
        
        For A = 1 To UserList(UserIndex).CurrentInventorySlots

            If .Invent.Object(A).ObjIndex > 0 Then
                Call QuitarUserInvItem(UserIndex, A, MAX_INVENTORY_OBJS)
                Call UpdateUserInv(False, UserIndex, A)
            End If

        Next A
        
        Call UpdateUserHechizos(True, UserIndex, 0)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
    End With

End Sub

' Otorgamos hechizos
Public Sub Event_UserSpells(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        .Stats.UserHechizos(1) = 20 ' Fuerza
        .Stats.UserHechizos(2) = 18 ' Agilidad
        
        .Stats.UserHechizos(1) = 10 ' Remover parálisis
        .Stats.UserHechizos(2) = 24 ' Inmovilizar
        .Stats.UserHechizos(3) = 61 ' Tormenta eléctrica
        .Stats.UserHechizos(4) = 62 ' Apocalipsis II
        
        Call UpdateUserHechizos(True, UserIndex, 0)
    End With

End Sub

' Inicia oficialmente el evento
Private Sub Event_Change(ByVal SlotEvent As Byte)

    Dim A As Long
    
    With Events(SlotEvent)
    
        For A = LBound(.Users) To UBound(.Users)

            If .Users(A).ID > 0 Then
                Call Event_ClassOld(SlotEvent, .Users(A).ID)
                Call Event_ChangeUser(.Users(A).ID, SlotEvent)
                Call EventWarpUser(.Users(A).ID, 65, 54, 22)
            End If

        Next A
        
    End With

End Sub

Public Sub CheckEvent_Time_Auto()

    On Error GoTo ErrHandler
    
    Dim A                            As Long

    Dim AllowedClass(1 To NUMCLASES) As Byte

    Dim AllowedFaction(1 To 4)       As Byte
    
    ' Clases válidas
    For A = 1 To NUMCLASES
        AllowedClass(A) = 1
    Next A
    
    ' Facciones válidas
    For A = 1 To 4
        AllowedFaction(A) = 1
    Next A
    
    ' Events automáticos
    Select Case Format(Now, "hh:mm")
        Case "13:00"
            NewEvent Busqueda, vbNullString, 10, 1, 47, 0, 0, 60, 360, 0, False, _
                0, 1, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
        Case "14:00"
            NewEvent DagaRusa, vbNullString, 10, 1, 47, 0, 0, 60, 360, 0, False, _
                0, 1, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
        Case "15:00"
            NewEvent GranBestia, vbNullString, 10, 1, 47, 0, 0, 60, 360, 0, False, _
                0, 1, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
        Case "16:00"
            NewEvent Teleports, vbNullString, 10, 1, 47, 0, 0, 60, 360, 0, False, _
                0, 1, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
        Case "17:00"
            NewEvent DeathMatch, vbNullString, 15, 1, 47, 0, 0, 60, 360, 0, False, _
                0, 2, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
        Case "18:00"
            NewEvent Unstoppable, vbNullString, 20, 1, 47, 0, 0, 60, 360, 0, False, _
                0, 2, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
    End Select
    
    If Hour(Now) >= 5 And Hour(Now) <= 9 Then Exit Sub
    
    Event_Time_Auto = Event_Time_Auto + 1
    
    
    
    Select Case Event_Time_Auto
        
        Case 15
            AllowedClass(eClass.Warrior) = 1
            AllowedClass(eClass.Hunter) = 1
            NewEvent Enfrentamientos, vbNullString, 8, 1, 47, 0, 0, 60, 180, 2, False, 1500, 1, 350000, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()

        Case 30
            AllowedClass(eClass.Hunter) = 0
            NewEvent Enfrentamientos, vbNullString, 8, 1, 47, 0, 0, 60, 180, 1, False, 1500, 3, 0, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()

        Case 60
            NewEvent Enfrentamientos, vbNullString, 12, 1, 47, 0, 0, 60, 180, 3, False, 1500, 1, 250000, 0, 0, False, False, 0, 0, 0, AllowedFaction(), AllowedClass()
            Event_Time_Auto = 0

        Case Else

            Exit Sub
        
    End Select
    
    Exit Sub

ErrHandler:
    Call LogEventos("Ocurrió un error en el CheckEvent_Time_Auto")
End Sub

Public Sub Event_RandomUsers_Array(ByVal Slot As Byte, ByRef vArray() As tUserEvent)
      
    Dim i          As Long

    Dim rndIndex   As Long

    Dim Temp       As tUserEvent

    Dim startIndex As Integer

    Dim endIndex   As Integer
    
    Randomize
      
    startIndex = LBound(vArray)
    endIndex = UBound(vArray)
      
    For i = startIndex To endIndex
        rndIndex = Int((endIndex - startIndex + 1) * Rnd() + startIndex)
  
        Temp = vArray(i)
        vArray(i) = vArray(rndIndex)
        vArray(rndIndex) = Temp
        
        With Events(Slot)

            If .Users(rndIndex).ID > 0 Then
                UserList(.Users(rndIndex).ID).flags.SlotUserEvent = rndIndex
            End If
            
            If .Users(i).ID > 0 Then
                UserList(.Users(i).ID).flags.SlotUserEvent = i
            End If

        End With

    Next i
    
End Sub

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
                    Temp = Left$(Temp, Len(Temp) - 2)
                    Temp = Temp & " y " & UserList(.Users(A).ID).Name & "."
                Else
                    Temp = Temp & UserList(.Users(A).ID).Name & ", "
                    
                End If
            End If

        Next A
        
        Event_Text_Users = Temp
        
    End With
    
    Exit Function

ErrHandler:
End Function

Public Function Event_Text_Users_VS(ByVal Slot As Byte) As String

    Dim A     As Long

    Dim TeamA As String

    Dim TeamB As String
    
    With Events(Slot)

        Select Case .Modality

            Case eModalityEvent.CastleMode

                For A = 1 To .Quotas

                    If .Users(A).ID > 0 Then
                        If A > (.Quotas / 2) Then
                            TeamA = TeamA & UserList(.Users(A).ID).Name & ", "
                        Else
                            TeamB = TeamB & UserList(.Users(A).ID).Name & ", "
                        End If
                    End If

                Next A
                
                TeamA = Left$(TeamA, Len(TeamA) - 2)
                TeamB = Left$(TeamB, Len(TeamB) - 2)
                Event_Text_Users_VS = TeamA & " VS " & TeamB
            
            Case eModalityEvent.DeathMatch

                For A = 1 To .Quotas
                    TeamA = TeamA & UserList(.Users(A).ID).Name & " VS "
                Next A
                
                TeamA = Left$(TeamA, Len(TeamA) - 4)
                Event_Text_Users_VS = TeamA
        End Select

    End With
    
End Function

Public Sub Eventos_Reset_All()
    Dim A As Long
    
    For A = 1 To MAX_EVENT_SIMULTANEO
        With Events(A)
            If .Run Then
                Call CloseEvent(A, , True)
            End If
        End With
    Next A
End Sub

'#################################
' EVENTO DE TELEPORTS
'#################################

Public Sub Events_Teleports_Init(ByVal SlotEvent As Byte)
    Const Respawn_Map As Byte = 65
    Const Respawn_X As Byte = 25
    Const Respawn_Y As Byte = 45
    
    Dim X As Long, Y As Long
    Dim A As Long, ElectionX As Long, Election As Boolean
    
    Y = 37
    
    ' Seteamos los Teleports a la posición de comienzo
    For A = 1 To 5
        For X = 18 To 33
            With MapData(Respawn_Map, X, Y)
                
                If .Blocked = 0 Then
                    .TileExit.Map = Respawn_Map
                    .TileExit.X = Respawn_X
                    .TileExit.Y = Respawn_Y
                End If
                
            End With
        Next X
        
        Y = Y - 5
    Next A
    
    ' Seteamos los Teleports al paso siguiente
    Y = 37
    For A = 1 To 5
        Do While Election = False
            Randomize
            ElectionX = RandomNumber(18, 33)
            
            With MapData(Respawn_Map, ElectionX, Y)
                
                If .Blocked = 0 Then
                    .TileExit.Map = Respawn_Map
                    .TileExit.X = 25
                    .TileExit.Y = Y - 2
                    Election = True
                End If

            End With
           
        Loop
        
        Y = Y - 5
        Election = False
    Next A
    
    With Events(SlotEvent)
        For A = LBound(.Users()) To UBound(.Users())
            If .Users(A).ID > 0 Then
                EventWarpUser .Users(A).ID, 65, 25, 45
            End If
        Next A
    End With
    
End Sub

Public Sub Events_Teleports_Finish(ByVal UserIndex As Integer)

    Dim strReward As String
    
    With UserList(UserIndex)
        Events(.flags.SlotEvent).TimeFinish = 3
        
        Call WriteConsoleMsg(UserIndex, "¡¡Felicitaciones!! Has ganado el Evento de Teleports.", FontTypeNames.FONTTYPE_INFOGREEN)
        
        Call PrizeUser(UserIndex, strReward)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.flags.SlotEvent, eModalityEvent.Teleports) & "» El personaje " & .Name & " ha encontrado el Teleport ganador. ¡¡Felicitaciones!!. " & strReward, FontTypeNames.FONTTYPE_GUILD))
    End With
End Sub



'#################################
' EVENTO DE GRANBESTIA
'#################################

Public Sub Events_GranBestia_Init(ByVal SlotEvent As Byte)
    
    Const Respawn_Map As Byte = 65
    Const Respawn_X_MIN As Byte = 48
    Const Respawn_X_MAX As Byte = 62
    
    Const Respawn_Y_MIN As Byte = 43
    Const Respawn_Y_MAX As Byte = 54
    
    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    Dim A As Long
    
    With Events(SlotEvent)
        
        Pos.Map = Respawn_Map
        Pos.X = 54
        Pos.Y = 48
        
        NpcIndex = SpawnNpc(NPC_GRAN_BESTIA, Pos, False, False)
        
        If NpcIndex = 0 Then
            Call CloseEvent(SlotEvent, strModality(SlotEvent, eModalityEvent.GranBestia) & "» Evento cancelado por respawn inválido de criatura", True)
        Else
            .NpcIndex = NpcIndex
            
            With Npclist(NpcIndex)
                .Stats.MaxHp = Events(SlotEvent).Quotas * 2000
                .Stats.MinHp = .Stats.MaxHp
            End With
            
            For A = LBound(.Users()) To UBound(.Users())
                If .Users(A).ID > 0 Then
                    EventWarpUser .Users(A).ID, Respawn_Map, RandomNumber(Respawn_X_MIN, Respawn_X_MAX), RandomNumber(Respawn_Y_MIN, Respawn_Y_MAX)
                End If
            Next A
            
        End If
    End With
    
End Sub

Public Sub Events_GranBestia_MuereNpc(ByVal UserIndex As Integer)
    
    Dim SlotEvent As Byte
    Dim strTemp(1) As String
    
    SlotEvent = UserList(UserIndex).flags.SlotEvent
    
    With Events(SlotEvent)
        Call PrizeUser(UserIndex, strTemp(1))
        strTemp(0) = strModality(SlotEvent, eModalityEvent.GranBestia) & "» El personaje " & UserList(UserIndex).Name & " ha acabado con la Gran Bestia. " & strTemp(1)
        
        Call CloseEvent(SlotEvent, strTemp(0))
    End With
End Sub

Public Sub Events_GranBestia_MuereUser(ByVal UserIndex As Integer)

    Dim SlotEvent As Byte
    Dim UserWinner As Integer
    Dim strTemp(1) As String
    
    SlotEvent = UserList(UserIndex).flags.SlotEvent
    
    With Events(SlotEvent)
        Call AbandonateEvent(UserIndex)
        
        If .Inscribed = 1 Then
            UserWinner = SearchLastUserEvent(SlotEvent)
            
            Call PrizeUser(UserWinner, strTemp(1))
            strTemp(0) = strModality(SlotEvent, eModalityEvent.GranBestia) & "» El personaje " & UserList(UserWinner).Name & " ha logrado sobrevivir a la Gran Bestia. " & strTemp(1)
        
            Call CloseEvent(SlotEvent, strTemp(0))
        End If
    
    End With
End Sub



' #### JUEGOS DEL HAMBRE
' ####
' ####

Public Sub JDH_Init(ByVal Slot As Byte)
On Error GoTo ErrHandler
    
    Dim A As Long
    Dim B As Long
    
    With Events(Slot)
        For A = LBound(.Users()) To UBound(.Users())

        ' Summon Chars
        If .Users(A).ID > 0 Then
            EventWarpUser .Users(A).ID, 64, 50, 86
        End If

        Next A
    
        
        ' Pociones de Energía y cofres
        For A = 1 To (.Quotas + 5)
            For B = 0 To 3
                NpcIndex = SpawnNpc(905 + B, Pos, False, False)
            Next B
            
            Event_MakeObj 629, 64, RandomNumber(25, 75), RandomNumber(15, 71), 0
        Next A
        
        
        
    End With
    
    
    Exit Sub
    
ErrHandler:
    LogEventos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : InitUnstoppable()"
End Sub
    
Public Sub Events_UserDie(ByVal UserIndex As Integer, _
                          ByVal AttackerIndex As Integer)
    
    Dim SlotEvent As Byte
    SlotEvent = .flags.SlotEvent
    
    With Events(SlotEvent)
        Select Case Events(SlotEvent).Modality
            Case eModalityEvent.CastleMode
                .Counters.TimeEventRevive = 3

            Case eModalityEvent.DeathMatch
                Call DeathMatch_UserDie(SlotEvent, UserIndex)

            Case eModalityEvent.Unstoppable
                .Users(UserList(AttackerIndex).flags.SlotUserEvent).Value = .Users(UserList(AttackerIndex).flags.SlotUserEvent).Value + 1
                Call EventWarpUser(UserIndex, 64, RandomNumber(30, 54), RandomNumber(25, 39))
                Call RevivirUsuario(UserIndex)

            Case eModalityEvent.Enfrentamientos
                Fight_UserDie SlotEvent, .flags.SlotUserEvent, AttackerIndex
        End Select
    
    
    End With
End Sub

Public Sub Events_UserRevive(ByVal UserIndex As Integer)
    
    Call RevivirUsuario(UserIndex)
    
    Call EventWarpUser(UserIndex, UserList(UserIndex).Pos.Map, RandomNumber(20, 70), RandomNumber(20, 70))
End Sub
