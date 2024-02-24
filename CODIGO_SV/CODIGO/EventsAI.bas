Attribute VB_Name = "EventsAI"
Option Explicit

Private Seed As Long

' @ Maximo de Mapas de Teleports
Public EVENTS_MAX_TELEPORTS As Byte

' @ Probabilidad de que el evento solicite que 'VALGA RESU'
Public EVENTS_PORC_RESUCITATION As Byte

' @ Probabilidad de que el evento solicite SIN 'ESCUDOS/CASCOS'
Public EVENTS_PORC_ESCUDO_CASCO As Byte

' @ Probabilidad de que el servidor determine el nivel que tendrán los personajes inscriptos.
Public EVENTS_PORC_LEVEL_CHANGE As Byte

' @ Index de configuracion de niveles para determinar en el cambio de clase y raza
Public EVENTS_SLOT_LEVEL_CHANGE As Byte

' @ Index de configuracion de niveles para determinar en el cambio de clase y raza
Public EVENTS_SLOT_LEVEL_CHANGE_POWER As Byte

' @ Probabilidad de que el evento cambie de CLASE-RAZA
Public EVENTS_PORC_CHANGE_CLASS_RAZE As Byte

' @ Probabilidad de que el evento permita 'OCULTAR-INVI'
Public EVENTS_PORC_VALE_OCULTAR_INVI As Byte

' @ Posibilidad de que sea PLANTES
Public EVENTS_PORC_PLANTE As Byte

' @ Posibilidad de que sea fuego amigo
Public EVENTS_PORC_FUEGOAMIGO As Byte

' @ Posibilidad de que sea con PARTY
Public EVENTS_PORC_PARTY As Byte


Public Const EVENTS_PORC_MITAD As Byte = 50

' @ Indice máximo de EVENTO (Ultima modalidad de evento)
Public EVENTS_INDICE_MAX_EVENT As Byte


Private Type tEvents_Level
    LvlMin As Byte
    LvlMax As Byte
End Type

Private Type tEvents_PorcMinMax
        Porc As Byte
        Min As Byte
        max As Byte
End Type

Private Type tEvents_Class
    Class(1 To NUMCLASES) As Byte
End Type

' @ PRE-Config de NIVELES
Private Const EVENTS_LEVEL_DEFAULT As Byte = 1
Private Events_Level_Last As Byte
Private Events_Level() As tEvents_Level

' @ PRE-Config de ROJAS
Private Events_Red_Last As Byte
Private Events_Red() As Integer

' @ PRE-Config de ORO
Private Events_Gold_Last As Byte
Private Events_Gold() As Integer

' @ PRE-Config de DSP
Private Events_Dsp_Last As Byte
Private Events_Dsp() As Integer

' @ PRE-Config de ROUNDS
Private Events_Rounds_Last As Byte
Private Events_Rounds() As Integer

' @ PRE-Config de ROUNDS FINALES
Private Events_Rounds_Final_Last As Byte
Private Events_Rounds_Final() As Integer

' @PRE-Config de Team Cant (VS)
Private Events_TeamCant_Last As Byte
Private Events_TeamCant() As tEvents_PorcMinMax

' @PRE-Config de Clases válidas
Private Events_Class_Last As Byte
Private Events_Class() As tEvents_Class
Private Const EVENTS_CLASS_TODAS_LUCHADORAS As Byte = 1
Private Const EVENTS_CLASS_NO_WAR_CAZA As Byte = 2
Private Const EVENTS_CLASS_MAGIC As Byte = 3
Private Const EVENTS_CLASS_SEMI_MAGIC As Byte = 4
Private Const EVENTS_CLASS_NO_MAGIC As Byte = 5
Private Const EVENTS_CLASS_TODAS As Byte = 6
Private Const EVENTS_CLASS_PLANTE As Byte = 7

' @ PRE-Config de OBJETOS para PREMIOS
Private Type tEventsObj
    ObjIndex As Integer         ' Objeto que otorga
    AmountMax As Integer        ' Máximo de cantidad que puede dar de ese item.
    Rank As Byte                ' Valor del 1-10 que determina que tan 'bueno' es el premio.
End Type

Private Events_Obj_Last As Integer
Private Events_Obj() As tEventsObj

Public Type tEvents_Automatic
    Events_Automatic_Active As Byte
    
    HourMin As String
    HourMax As String
    SecondDelay As Long
End Type

Public Events_Automatic As tEvents_Automatic

' @ Eventos que hace JARVIS
Public Sub Events_Automatic_Loop()
        '<EhHeader>
        On Error GoTo Events_Automatic_Loop_Err
        '</EhHeader>
        
        Static Second As Integer
    
100     Second = Second + 1
    
102     If Events_Automatic.Events_Automatic_Active = 0 Then
104         Second = 0
            Exit Sub
        End If
    
        Dim Time As Date
106     Time = Format(Now, "hh:mm")
    
108     If Not (Time >= Events_Automatic.HourMin Or Time <= Events_Automatic.HourMax) Then Exit Sub

110     If Second >= Events_Automatic.SecondDelay Then
112         Call Events_SetConfig
114         Second = 0
        End If
    
        '<EhFooter>
        Exit Sub

Events_Automatic_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventsAI.Events_Automatic_Loop " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' # Busca un Mapa disponible según requisitos
Public Function Events_Teleports_SearchMapFree(ByVal Quotas As Byte)

On Error GoTo ErrHandler

    Dim A As Long
    
    
    For A = 1 To EVENTS_MAX_TELEPORTS
        With MapEvent.Teleport(A)
            If .MaxQuotas >= Quotas And .Usage = False Then
                Events_Teleports_SearchMapFree = A
                
                Exit Function
            End If
        
        End With
    Next A
    
    Exit Function
ErrHandler:
End Function
' # Cargamos la configuracion de los portales
Public Sub Events_Load_TeleportsMap(ByRef Manager As clsIniManager)

    On Error GoTo ErrHandler
    
    Dim A As Long
    Dim Temp As String
    
    EVENTS_MAX_TELEPORTS = val(Manager.GetValue("CONST", "EVENTS_MAX_TELEPORTS"))
    
    ReDim MapEvent.Teleport(1 To EVENTS_MAX_TELEPORTS) As tMapTeleport
    
    For A = 1 To EVENTS_MAX_TELEPORTS
        With MapEvent.Teleport(A)
            Temp = Manager.GetValue("TELEPORT" & A, "Map")
            
            .Map = val(ReadField(1, Temp, 45))
            .XWarp = val(ReadField(2, Temp, 45))
            .YWarp = val(ReadField(3, Temp, 45))
            
            
            .XInitial_TP = Manager.GetValue("TELEPORT" & A, "XInitial_TP")
            
            .XTiles_TP = Manager.GetValue("TELEPORT" & A, "XTiles_TP")
            .Y_Pasajes = Manager.GetValue("TELEPORT" & A, "Y_Pasajes")
            .Y_TileAdd = Manager.GetValue("TELEPORT" & A, "YAdd_TP")
            .MaxQuotas = Manager.GetValue("TELEPORT" & A, "XTiles_TP")
        End With
    Next A
    
    Exit Sub
ErrHandler:
    
End Sub
Public Sub Events_Load_PreConfig()
        '<EhHeader>
        On Error GoTo Events_Load_PreConfig_Err
        '</EhHeader>
        Dim Manager As clsIniManager
        Dim FilePath As String
        Dim Temp As String
        Dim A As Long, B As Long
        Dim ArraiByte() As String
    
100     Set Manager = New clsIniManager
    
102     FilePath = DatPath & "events\events_config.ini"
104     Manager.Initialize FilePath
    
    
106     Events_Automatic.HourMin = Manager.GetValue("HOURS", "HOURMIN")
108     Events_Automatic.HourMax = Manager.GetValue("HOURS", "HOURMAX")
110     Events_Automatic.SecondDelay = val(Manager.GetValue("HOURS", "SECONDS_DELAY"))
    
    
        ' Mapas de Teleports
        
        
        ' Constantes de PORCENTAJE
112     EVENTS_PORC_RESUCITATION = val(Manager.GetValue("CONST", "EVENTS_PORC_RESUCITATION"))
114     EVENTS_PORC_ESCUDO_CASCO = val(Manager.GetValue("CONST", "EVENTS_PORC_ESCUDO_CASCO"))
116     EVENTS_PORC_LEVEL_CHANGE = val(Manager.GetValue("CONST", "EVENTS_PORC_LEVEL_CHANGE"))
118     EVENTS_INDICE_MAX_EVENT = val(Manager.GetValue("CONST", "EVENTS_INDICE_MAX_EVENT"))
120     EVENTS_PORC_CHANGE_CLASS_RAZE = val(Manager.GetValue("CONST", "EVENTS_PORC_CHANGE_CLASS_RAZE"))
122     EVENTS_PORC_VALE_OCULTAR_INVI = val(Manager.GetValue("CONST", "EVENTS_PORC_VALE_OCULTAR_INVI"))
          EVENTS_PORC_PLANTE = val(Manager.GetValue("CONST", "EVENTS_PORC_PLANTE"))
          EVENTS_PORC_FUEGOAMIGO = val(Manager.GetValue("CONST", "EVENTS_PORC_FUEGOAMIGO"))
          EVENTS_PORC_PARTY = val(Manager.GetValue("CONST", "EVENTS_PORC_PARTY"))
          
        ' Comprobaciones de NIVEL
124     Events_Level_Last = val(Manager.GetValue("LEVEL", "LAST"))
          EVENTS_SLOT_LEVEL_CHANGE = val(Manager.GetValue("LEVEL", "INDEXCHANGE"))
          EVENTS_SLOT_LEVEL_CHANGE_POWER = val(Manager.GetValue("LEVEL", "INDEXCHANGEPOWER"))
          
126     ReDim Events_Level(1 To Events_Level_Last) As tEvents_Level
    
128     For A = 1 To Events_Level_Last
130         Temp = Manager.GetValue("LEVEL", CStr(A))
132         Events_Level(A).LvlMin = val(ReadField(1, Temp, 45))
134         Events_Level(A).LvlMax = val(ReadField(2, Temp, 45))
136     Next A
    
        ' Comprobaciones de ROJAS
138     Events_Red_Last = val(Manager.GetValue("RED", "LAST"))
    
140     ReDim Events_Red(1 To Events_Red_Last) As Integer
    
142     For A = 1 To Events_Red_Last
144          Events_Red(A) = val(Manager.GetValue("RED", CStr(A)))
146     Next A
        
        ' Combinaciones de ORO
         Events_Gold_Last = val(Manager.GetValue("GOLD", "LAST"))
        
         ReDim Events_Gold(1 To Events_Gold_Last) As Integer
        
         For A = 1 To Events_Gold_Last
              Events_Gold(A) = val(Manager.GetValue("GOLD", CStr(A)))
         Next A
            
        ' Combinaciones de DSP
         Events_Dsp_Last = val(Manager.GetValue("DSP", "LAST"))
        
         ReDim Events_Dsp(1 To Events_Dsp_Last) As Integer
        
         For A = 1 To Events_Dsp_Last
              Events_Dsp(A) = val(Manager.GetValue("DSP", CStr(A)))
         Next A
        
        
        ' Comprobaciones de ROUNDS
148     Events_Rounds_Last = val(Manager.GetValue("ROUNDS", "LAST"))
    
150     ReDim Events_Rounds(1 To Events_Rounds_Last) As Integer
    
152     For A = 1 To Events_Rounds_Last
154          Events_Rounds(A) = val(Manager.GetValue("ROUNDS", CStr(A)))
156     Next A
    
        ' Comprobaciones de ROUNDS FINALES
158     Events_Rounds_Final_Last = val(Manager.GetValue("ROUNDS_FINAL", "LAST"))
    
160     ReDim Events_Rounds_Final(1 To Events_Rounds_Final_Last) As Integer
    
162     For A = 1 To Events_Rounds_Final_Last
164          Events_Rounds_Final(A) = val(Manager.GetValue("ROUNDS_FINAL", CStr(A)))
166     Next A
    
        ' Comprobaciones de TEAMCANT
168     Events_TeamCant_Last = val(Manager.GetValue("TEAMCANT", "LAST"))
    
170     ReDim Events_TeamCant(1 To Events_TeamCant_Last) As tEvents_PorcMinMax
    
172     For A = 1 To Events_TeamCant_Last
174         Temp = Manager.GetValue("TEAMCANT", CStr(A))
        
176         Events_TeamCant(A).Porc = val(ReadField(1, Temp, 45))
178         Events_TeamCant(A).Min = val(ReadField(2, Temp, 45))
180         Events_TeamCant(A).max = val(ReadField(3, Temp, 45))
182     Next A
    
        ' Comprobaciones de CLASES
184     Events_Class_Last = val(Manager.GetValue("CLASS", "LAST"))
    
186     ReDim Events_Class(1 To Events_Class_Last) As tEvents_Class
    
188     For A = 1 To Events_Class_Last
190         Temp = Manager.GetValue("CLASS", CStr(A))
        
192         ArraiByte = Split(Temp, "-")
        
194         For B = LBound(ArraiByte) To UBound(ArraiByte)
196             Events_Class(A).Class(B + 1) = val(ArraiByte(B))
198         Next B
200     Next A
        
        ' # Objetos de PREMIOS
        Events_Obj_Last = val(Manager.GetValue("PRIZE_OBJ", "LAST"))
        
        ReDim Events_Obj(1 To Events_Obj_Last) As tEventsObj
        
        For A = 1 To Events_Obj_Last
            Temp = Manager.GetValue("PRIZE_OBJ", "Obj" & A)
        
            Events_Obj(A).ObjIndex = val(ReadField(1, Temp, 45))
            Events_Obj(A).AmountMax = val(ReadField(2, Temp, 45))
            Events_Obj(A).Rank = val(ReadField(3, Temp, 45))

        Next A
        
202     Set Manager = Nothing
    
        '<EhFooter>
        Exit Sub

Events_Load_PreConfig_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventsAI.Events_Load_PreConfig " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function Events_SetMaxQuote(ByRef TempEvent As tEvents)

    On Error GoTo ErrHandler
    
    With TempEvent

        Select Case .Modality
    
            Case eModalityEvent.Enfrentamientos

                    
                Select Case .TeamCant
                    
                    Case 1
                        .QuotasMax = 64
                        .QuotasMin = 2
                    
                    Case 2
                        .QuotasMax = 64
                        .QuotasMin = (.TeamCant * 4)


                    Case 4, 8
                        .QuotasMax = 64
                        .QuotasMin = (.TeamCant * 2)
                        
                    Case 3, 6
                        .QuotasMax = 48
                         .QuotasMin = (.TeamCant * 2)
                         
                    Case 5, 10
                        .QuotasMax = 80
                        .QuotasMin = (.TeamCant * 2)
                        
                    Case 7
                        .QuotasMax = 56
                         .QuotasMin = (.TeamCant * 2)

                    Case 9
                        .QuotasMax = 72
                         .QuotasMin = (.TeamCant * 4)
                End Select
                    
            Case eModalityEvent.Busqueda
                .QuotasMin = 10
                .QuotasMax = 20
                
            Case eModalityEvent.CastleMode
                .QuotasMin = 10
                .QuotasMax = 30
                
            Case eModalityEvent.DagaRusa
                .QuotasMin = 5
                .QuotasMax = 30
                
            Case eModalityEvent.DeathMatch
                .QuotasMin = 7
                .QuotasMax = 20
                
            Case eModalityEvent.GranBestia
                .QuotasMin = 5
                .QuotasMax = 10
                
            Case eModalityEvent.JuegosDelHambre
                .QuotasMin = 30
                .QuotasMax = 30
                
            Case eModalityEvent.Teleports
                .QuotasMin = 5
                .QuotasMax = 60
                
            Case eModalityEvent.Unstoppable
                .QuotasMin = 7
                .QuotasMax = 30
                
        End Select
    
    End With
    
    Exit Function
ErrHandler:
   
End Function

' Setea el nivel mínimo requerido o el nivel cambiado
Private Function Events_SetLevel(ByRef TempEvent As tEvents)
    
    On Error GoTo ErrHandler
    
    Dim RandomLevel       As Byte

    Dim RandomLevelServer As Boolean
        
    With TempEvent
        
        Select Case .Modality
        
            Case eModalityEvent.Enfrentamientos, eModalityEvent.DeathMatch, eModalityEvent.Unstoppable, eModalityEvent.CastleMode, eModalityEvent.GranBestia
                        
                
                
                If .ChangeClass > 0 Or .ChangeRaze > 0 Then
                    RandomLevelServer = (RandomNumberPower(1, 100) <= EVENTS_PORC_LEVEL_CHANGE)
                    
                    If RandomLevelServer Then
                        .LvlMin = Events_Level(EVENTS_LEVEL_DEFAULT).LvlMin
                        .LvlMax = Events_Level(EVENTS_LEVEL_DEFAULT).LvlMax
                        
                        
                        If RandomNumberPower(1, 100) <= EVENTS_PORC_MITAD Then
                            RandomLevel = RandomNumberPower(EVENTS_SLOT_LEVEL_CHANGE + 1, Events_Level_Last) ' 47-50, 50-50, 60-60, 70-70
                        Else
                            RandomLevel = EVENTS_SLOT_LEVEL_CHANGE '40 a 47
                            
                        End If
                        
                         .ChangeLevel = RandomNumberPower(Events_Level(RandomLevel).LvlMin, Events_Level(RandomLevel).LvlMax)
                        Exit Function

                    End If
                    
                End If
                
                
                RandomLevel = RandomNumberPower(EVENTS_LEVEL_DEFAULT + 1, EVENTS_SLOT_LEVEL_CHANGE - 1)
                    
                .LvlMin = Events_Level(RandomLevel).LvlMin
                .LvlMax = Events_Level(RandomLevel).LvlMax

            Case Else
                .LvlMin = Events_Level(EVENTS_LEVEL_DEFAULT).LvlMin
                .LvlMax = Events_Level(EVENTS_LEVEL_DEFAULT).LvlMax
            
        End Select
    
    End With
    
    Exit Function
    
ErrHandler:
End Function

Private Function Events_Modify_Points_Rounds(ByVal RoundsFinal As Byte) As Single
        Events_Modify_Points_Rounds = 1
End Function
Public Function Events_SetReward_Points(ByRef TempEvent As tEvents, ByVal Quotas As Integer) As Integer
        '<EhHeader>
        On Error GoTo Events_SetReward_Points_Err
        '</EhHeader>
        
        Dim Bonus_Rounds As Long
    
100     With TempEvent
    
102         Select Case .Modality
        
                Case eModalityEvent.Enfrentamientos
                    Events_SetReward_Points = (Quotas * (1 / .TeamCant))
                    Bonus_Rounds = .LimitRound * ((.LimitRoundFinal + 1) ^ 0.6) '
                    
                    If TempEvent.config(eConfigEvent.eParty) = 1 Then
                        Events_SetReward_Points = Events_SetReward_Points * 1.5
                    End If

110             Case eModalityEvent.DagaRusa, eModalityEvent.Teleports
112                 Events_SetReward_Points = (Quotas / 2)
                
114             Case eModalityEvent.DeathMatch
116                 Events_SetReward_Points = (Quotas - .TeamCant) / 3
                
118             Case eModalityEvent.Unstoppable
120                 Events_SetReward_Points = (Quotas - .TeamCant) / 2
                
122             Case eModalityEvent.CastleMode
124                 Events_SetReward_Points = (Quotas) / 2
                
126             Case Else
128                 Events_SetReward_Points = 1

            End Select
        
        
130         Events_SetReward_Points = (Events_SetReward_Points + Bonus_Rounds)
        End With

        '<EhFooter>
        Exit Function

Events_SetReward_Points_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventsAI.Events_SetReward_Points " & _
               "at line " & Erl
        '</EhFooter>
End Function
Public Function Events_SetInscription_Gold(ByRef TempEvent As tEvents, ByVal Quotas As Integer) As Long
    On Error GoTo Events_SetInscription_Gold_Err
    
    Dim Bonus_Rounds As Long
    Dim Valor As Long
    Dim ValorBase As Long
    Dim Multiplicador As Double

    With TempEvent
        Select Case .Modality
            Case eModalityEvent.Enfrentamientos
                ValorBase = RandomNumber(1, Events_Gold_Last)
                ValorBase = Events_Gold(ValorBase)
                Valor = ValorBase
                
                Events_SetInscription_Gold = (Quotas * (1 / .TeamCant))
                Bonus_Rounds = .LimitRound * ((.LimitRoundFinal + 1))

                If TempEvent.config(eConfigEvent.eParty) = 1 Then
                    Events_SetInscription_Gold = Events_SetInscription_Gold * 1.5
                End If


            Case Else
                Events_SetInscription_Gold = 1 ' Valor predeterminado si no coincide con ninguna modalidad específica
        End Select

        Events_SetInscription_Gold = Valor * (Events_SetInscription_Gold + Bonus_Rounds)
    End With

    Exit Function

Events_SetInscription_Gold_Err:
    LogError Err.description & vbCrLf & _
           "in ServidorArgentum.EventsAI.Events_SetInscription_Gold " & _
           "at line " & Erl
End Function

Private Function RandomNormal(Mean As Double, StdDev As Double) As Double
    ' Genera un número aleatorio con distribución normal
    RandomNormal = Mean + (StdDev * Rnd)
End Function
Private Sub Events_SetClass_Change(ByRef TempEvent As tEvents)
    On Error GoTo Events_With_Party_Err
    
    With TempEvent
        
        If .TeamCant > 1 Then

            ' @ [Parejas con PARTY]
            If .TeamCant <= 5 Then
                If RandomNumberPower(1, 100) <= (EVENTS_PORC_PARTY - (10 * .TeamCant)) Then
                    .config(eConfigEvent.eParty) = 1
                End If
            End If
            
            ' Posibilidad de resucitar a los compañeros
            If .ChangeLevel >= 40 And (.ChangeClass <> eClass.Paladin And .ChangeClass <> eClass.Hunter And .ChangeClass <> eClass.Paladin And .ChangeClass <> eClass.Assasin) Then
                .config(eConfigEvent.eResu) = IIf(RandomNumberPower(1, 100) <= EVENTS_PORC_RESUCITATION, 1, 0)
            End If
            
            ' Posibilidad de poder matar a tus compañeros
            .config(eConfigEvent.eFuegoAmigo) = IIf(RandomNumberPower(1, 100) <= (EVENTS_PORC_FUEGOAMIGO), 1, 0)
        Else
        
            If .ChangeClass = eClass.Warrior Or .ChangeClass = eClass.Hunter Or .ChangeClass = eClass.Assasin Or .ChangeClass = eClass.Paladin Then
             
                If .ChangeClass = eClass.Warrior Or .ChangeClass = eClass.Hunter Then
                    .IsPlante = 1
                Else

                    If RandomNumberPower(1, 100) <= EVENTS_PORC_PLANTE Then
                        .IsPlante = 1
                    End If
                    
                End If
              
                If RandomNumberPower(1, 100) <= EVENTS_PORC_ESCUDO_CASCO Then
                    .config(eConfigEvent.eCascoEscudo) = 0
                End If
                
            End If
        End If

    End With

    '<EhFooter>
    Exit Sub

Events_With_Party_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.EventsAI.Events_With_Party " & "at line " & Erl
  
    '</EhFooter>
End Sub

Private Sub Events_SetClass_NotChange(ByRef TempEvent As tEvents)
        '<EhHeader>
        On Error GoTo Events_SetClass_Valid_Err
        '</EhHeader>

        Dim A           As Long

        Dim RandomClass As Byte
    
100     With TempEvent

            If .TeamCant > 1 Then
                
                ' Posibilidad de poder matar a tus compañeros
                TempEvent.config(eConfigEvent.eFuegoAmigo) = IIf(RandomNumberPower(1, 100) <= (EVENTS_PORC_FUEGOAMIGO), 1, 0)
            
                ' @ Prob de dar resu
                If .LvlMin >= 40 Then
                    .config(eConfigEvent.eResu) = IIf(RandomNumberPower(1, 100) <= EVENTS_PORC_RESUCITATION, 1, 0)
                End If
                
                ' @ [Parejas con PARTY]
                If .TeamCant <= 5 Then
                    If RandomNumberPower(1, 100) <= (EVENTS_PORC_PARTY - (10 * .TeamCant)) Then
                        .config(eConfigEvent.eParty) = 1
                    End If
                End If

                ' @ 2vs2: NO queremos GUERREROS-CAZADORES en el equipo AZAR.
102             If TempEvent.TeamCant = 2 Then
104                 If TempEvent.config(eConfigEvent.eParty) = 0 Then
106                     .AllowedClasses = Events_Class(EVENTS_CLASS_NO_WAR_CAZA).Class
                    Else
108                     .AllowedClasses = Events_Class(EVENTS_CLASS_TODAS_LUCHADORAS).Class
                    End If

                Else
110                 .AllowedClasses = Events_Class(EVENTS_CLASS_TODAS_LUCHADORAS).Class
                End If

            Else
            
                ' 20% de que elija una clase específica
112             If RandomNumberPower(1, 100) <= 20 Then
114                 RandomClass = RandomNumberPower(1, NUMCLASES - 2)        ' NO work, thief
                
116                 ReDim .AllowedClasses(1 To NUMCLASES) As Byte

118                 For A = 1 To NUMCLASES
120                     .AllowedClasses(A) = 0
122                 Next A
                        
124                 .AllowedClasses(RandomClass) = 1
                         
                    ' GUERRERO-CAZADOR plantan
126                 If RandomClass = eClass.Warrior Or RandomClass = eClass.Hunter Then
128                     .IsPlante = 1
                    
130                     If RandomNumberPower(1, 100) <= EVENTS_PORC_ESCUDO_CASCO Then
132                         .config(eConfigEvent.eCascoEscudo) = 0
                        End If

134                 ElseIf RandomClass = eClass.Paladin Or eClass.Assasin Then

136                     If RandomNumberPower(1, 100) <= EVENTS_PORC_PLANTE Then
138                         .IsPlante = 1

                        End If

                    End If

                Else
    
140                 If RandomNumberPower(1, 100) <= 50 Then
142                     If RandomNumberPower(1, 100) <= 50 Then
144                         If RandomNumberPower(1, 100) <= 50 Then
146                             .AllowedClasses = Events_Class(EVENTS_CLASS_SEMI_MAGIC).Class
                            Else
148                             .AllowedClasses = Events_Class(EVENTS_CLASS_PLANTE).Class
150                             .IsPlante = 1
                            End If
                             
152                         If RandomNumberPower(1, 100) <= EVENTS_PORC_PLANTE Then
154                             .IsPlante = 1
                            End If
                            
156                         If RandomNumberPower(1, 100) <= EVENTS_PORC_ESCUDO_CASCO Then
158                             .config(eConfigEvent.eCascoEscudo) = 0
                            End If

                        Else

160                         If RandomNumberPower(1, 100) <= 30 Then
162                             .AllowedClasses = Events_Class(EVENTS_CLASS_NO_MAGIC).Class
                            Else
164                             .AllowedClasses = Events_Class(EVENTS_CLASS_PLANTE).Class
                            End If
                              
166                         .IsPlante = 1
                            
168                         If RandomNumberPower(1, 100) <= EVENTS_PORC_ESCUDO_CASCO Then
170                             .config(eConfigEvent.eCascoEscudo) = 0
                            End If

                        End If

                    Else
172                     .AllowedClasses = Events_Class(EVENTS_CLASS_MAGIC).Class

                    End If

                End If

            End If
    
        End With

        '<EhFooter>
        Exit Sub

Events_SetClass_Valid_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.EventsAI.Events_SetClass_Valid " & "at line " & Erl
  
        '</EhFooter>
End Sub

Private Sub Events_SetTeamCant(ByRef TempEvent As tEvents)
        '<EhHeader>
        On Error GoTo Events_SetTeamCant_Err
        '</EhHeader>
        
        Dim A       As Long
        Dim RandomB As Byte
    
100     With TempEvent
        
            ' 100%  Cambio de clase a (GUERREROS o CAZADORES) ¡PLANTAN SIEMPRE! por lo cual es 1.
102         If (.ChangeClass = eClass.Warrior Or .ChangeClass = eClass.Hunter) Then
104             TempEvent.TeamCant = 1
                Exit Sub
            End If
        
106         For A = 1 To Events_TeamCant_Last
108             RandomB = RandomNumberPower(1, 100)
            
110             If RandomB <= Events_TeamCant(A).Porc Then
112                 TempEvent.TeamCant = RandomNumberPower(Events_TeamCant(A).Min, Events_TeamCant(A).max)
                    Exit Sub
                End If

114         Next A
        
116         TempEvent.TeamCant = RandomNumberPower(1, 3)

            
            
        End With

        '<EhFooter>
        Exit Sub

Events_SetTeamCant_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.EventsAI.Events_SetTeamCant " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Events_SetRed(ByRef TempEvent As tEvents, ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo Events_SetRed_Err

        '</EhHeader>

        Dim LimitRed As Long

        Dim Obj      As Obj

        'Dim A        As Long
    
100     With TempEvent

102         If .LimitRed > 0 Then
104             LimitRed = CLng((.LimitRed) * (.LimitRound + .LimitRoundFinal))
            
106             If LimitRed > MAX_INVENTORY_OBJS Then LimitRed = MAX_INVENTORY_OBJS
            
110             Obj.Amount = LimitRed
112             Obj.ObjIndex = POCION_ROJA
            
114             'For A = LBound(.Users) To UBound(.Users)
118             Call QuitarObjetos(POCION_ROJA, MAX_INVENTORY_OBJS, UserIndex)
                    
120             If Not MeterItemEnInventario(UserIndex, Obj) Then
122                 Call LogError("El personaje " & UserList(UserIndex).Name & " no recibió las pociones rojas por falta de espacio.")

                End If

124             ' Next A
            
            End If

        End With

        '<EhFooter>
        Exit Sub

Events_SetRed_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.EventsAI.Events_SetRed " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

Public Sub Events_SetConfig()
        '<EhHeader>
        On Error GoTo Events_SetConfig_Err
        '</EhHeader>

        Dim TempEvent       As tEvents

        Dim ChangeClassRaze As Boolean
    
100     With TempEvent
            
            Dim Rano As Byte
            Rano = 100
            
            ' @ Determina la MODALIDAD del evento
            If RandomNumberPower(1, 100) <= Rano Then
                .Modality = eModalityEvent.Enfrentamientos
            Else
102             .Modality = eModalityEvent.Teleports

            End If
             
            ' Configuración INICIAL: VALE ESCUDOS
            .config(eConfigEvent.eCascoEscudo) = 1
            
            ' @ Determina si es cambio de CLASE-RAZA
            ' @ Determina las ROJAS mínima.
            ' @ Determina los ROUNDS.
104         If .Modality = eModalityEvent.Enfrentamientos Then
106             ChangeClassRaze = RandomNumberPower(1, 100) <= EVENTS_PORC_CHANGE_CLASS_RAZE
                    
108             If ChangeClassRaze Then
110                 .ChangeClass = RandomNumberPower(1, NUMCLASES - 2) ' Saco las clases TRABAJADOR-LADRÓN momentaneamente.
112                 .ChangeRaze = RandomNumberPower(1, NUMRAZAS)
                    .config(eConfigEvent.eInvFree) = 1
                      
114                 If .ChangeClass = eClass.Thief Then
116                     .config(eConfigEvent.eOcultar) = IIf(RandomNumberPower(1, 100) <= EVENTS_PORC_VALE_OCULTAR_INVI, 1, 0)
                    End If
                        
118                 .AllowedClasses = Events_Class(EVENTS_CLASS_TODAS).Class
                    
                    If .ChangeClass = eClass.Paladin Or .ChangeClass = eClass.Assasin Or .ChangeClass = eClass.Warrior Or .ChangeClass = eClass.Hunter Then
                        If RandomNumberPower(1, 100) <= EVENTS_PORC_ESCUDO_CASCO Then
                            .config(eConfigEvent.eCascoEscudo) = 0
                        End If
                    End If
                        
                End If
                    
120             .LimitRed = Events_Red(RandomNumberPower(1, Events_Red_Last))
122             .LimitRound = Events_Rounds(RandomNumberPower(1, Events_Rounds_Last))
124             .LimitRoundFinal = Events_Rounds_Final(RandomNumberPower(1, Events_Rounds_Final_Last))
                
            End If
        
            ' @ Nivel que permitirá ingresar al EVENTO.
126         Call Events_SetLevel(TempEvent)
            
           
            'TempEvent.TeamCant = 1 ' @ Only 1vs1
            
128         Select Case .Modality
                
                Case eModalityEvent.Enfrentamientos
                    
130                 Call Events_SetTeamCant(TempEvent)

                    'TempEvent.TeamCant = 1 ' @ Only 1vs1
                    
132                 If ChangeClassRaze Then
                        Call Events_SetClass_Change(TempEvent)
                    Else
                        Call Events_SetClass_NotChange(TempEvent)
                    End If
                    
                    
                    
                
140             Case eModalityEvent.DeathMatch, eModalityEvent.Unstoppable
142                 .AllowedClasses = Events_Class(EVENTS_CLASS_NO_WAR_CAZA).Class
                
144             Case Else
146                 .AllowedClasses = Events_Class(EVENTS_CLASS_TODAS).Class

            End Select
            
            ' Ajustamos el PREMIOSKI
            Call Events_Set_Prize_Obj(TempEvent)
            
            ' Cupos máximos posibles
148         Call Events_SetMaxQuote(TempEvent)
        
            ' Arenas disponibles
150         Call Events_SetArenas(TempEvent)
        
            ' Premio máximo que podrá ganar.
152         .PrizePoints = Events_SetReward_Points(TempEvent, .QuotasMax)
            .InscriptionGld = Events_SetInscription_Gold(TempEvent, .QuotasMax)
            '.PrizeEldhir = Events_SetReward_Dsp(TempEvent, .QuotasMax)
            
154         .Name = MODALITY_STRING(.Modality, .TeamCant, .IsPlante)

156         ReDim .AllowedFaction(1 To 4) As Byte
            
158         .Prob = RandomNumberPower(1, 100)
        
160         .AllowedFaction(1) = 1
162         .AllowedFaction(2) = 1
164         .AllowedFaction(3) = 1
166         .AllowedFaction(4) = 1
168         .config(eConfigEvent.eBronce) = 1
170         .config(eConfigEvent.ePlata) = 1
172         .config(eConfigEvent.eOro) = 1
174         .config(eConfigEvent.ePremium) = 1
175         .config(eConfigEvent.eUseApocalipsis) = 1
            .config(eConfigEvent.eUseDescarga) = 1
            .config(eConfigEvent.eUseParalizar) = 1
            .config(eConfigEvent.eUsePotion) = 1
            .config(eConfigEvent.eUseTormenta) = 1
              
176         .TimeCancel = 300
            
            Dim CanEvent As Byte
        
180         CanEvent = NewEvent(TempEvent, , "JARVIS")
        
182         If CanEvent <> 0 Then
184             Events(CanEvent).Enabled = True
            End If
        
        End With

        '<EhFooter>
        Exit Sub

Events_SetConfig_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.EventsAI.Events_SetConfig " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Events_SetArenas(ByRef TempEvent As tEvents)
        '<EhHeader>
        On Error GoTo Events_SetArenas_Err
        '</EhHeader>

        Dim A As Long
    
100     With TempEvent
102         .ArenasLimit = 4        ' 4 Retos en curso. VER
        
104         If .IsPlante > 0 Then
106             .ArenasMin = MAX_MAP_FIGHT_NORMAL + 1
108             .ArenasMax = MAX_MAP_FIGHT_PLANTES
            Else
110             .ArenasMin = 1
112             .ArenasMax = MAX_MAP_FIGHT_NORMAL
            End If
        
        End With

        '<EhFooter>
        Exit Sub

Events_SetArenas_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventsAI.Events_SetArenas " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' # Determina el PREMIO del evento.
Private Sub Events_Set_Prize_Obj(ByRef TempEvent As tEvents)
    
    
    On Error GoTo ErrHandler
    
    Dim MAX_CLASS As Byte
    Dim A As Long
    
    
    If Events_Obj_Last = 0 Then Exit Sub
    With TempEvent

        Dim ObjIndex As Integer
        Dim Random As Byte
        
        Random = RandomNumber(1, Events_Obj_Last)
        
        ObjIndex = Events_Obj(Random).ObjIndex            ' # Generado de manera aleatorea de la lista
        
        ' # Comprueba si las clases del torneo válidas utilizan el objeto que salió al azar
        ' # Esto hace que el torneo tengas más chances de que NO de muchos PREMIOS específicos para clase.
        ' # Hacer Excepción con MD
        
        If Not ObjIndex = EspadaMataDragonesIndex Then
            For A = 1 To NUMCLASES
                If .AllowedClasses(A) = 1 Then
                    If Not ClasePuedeItem(A, ObjIndex) Then
                        Exit Sub
                    End If
                End If
            Next A
        End If
        
        
        ' # Asigna el PREMIO al TORNEO
        
        .PrizeObj.ObjIndex = ObjIndex
        .PrizeObj.Amount = RandomNumber(1, Events_Obj(Random).AmountMax)
    End With
    
    Exit Sub
ErrHandler:
    
End Sub
