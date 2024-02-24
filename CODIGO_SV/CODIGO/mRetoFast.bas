Attribute VB_Name = "mRetoFast"
Option Explicit

Public Const RETOFAST_POINTS As Long = 0
Public Const RETOFAST_GOLD As Long = 0
Public Const MAX_RETO_FAST As Byte = 30
Public Const MAX_RETO_FAST_POSITIONS As Byte = 9


Public Enum eConfigVale
    ConfigDefault = 0
    ValeResu = 1        ' VALE RESU. SIN LIMITE DE ROJAS
    ValeTodo = 2        ' VALE TODO . SIN LIMITE DE ROJAS
End Enum

Private Type tMap

    X As Byte
    Y As Byte

End Type

Private Type tRetoFastConfig

    Ident As String
    Map As Integer
    Pos() As tMap
    Users() As Integer
    ConfigVale As eConfigVale
    
    RedLimit As Integer
End Type

Private Type tRetoFast
    Ident As String
    Run As Boolean
    Users() As Integer
    ArenaIndex As Integer
    ConfigVale As eConfigVale
    
    RedLimit As Integer
End Type

Public RetoFastConfig(1 To MAX_RETO_FAST_POSITIONS) As tRetoFastConfig
Public RetoFast(1 To MAX_RETO_FAST) As tRetoFast

' # Cargamos la información de los Retos rápidos
Public Sub LoadRetoFast()
    
    ' # 1vs1 Default
    With RetoFastConfig(1)
        .Ident = "1vs1"
        .Map = 1
        
        ReDim .Pos(1 To 2) As tMap
        ReDim .Users(1 To 2) As Integer
        
        .Pos(1).X = 71
        .Pos(1).Y = 49
        
        .Pos(2).X = 72
        .Pos(2).Y = 49
        
    End With
        
    
    ' # 2vs2 Default
    With RetoFastConfig(2)
        .Ident = "2vs2"
        ReDim .Pos(1 To 4) As tMap
        ReDim .Users(1 To 4) As Integer
        
        .Map = 1
        .Pos(1).X = 71
        .Pos(1).Y = 52
        .Pos(2).X = 71
        .Pos(2).Y = 53
        
        .Pos(3).X = 72
        .Pos(3).Y = 52
        .Pos(4).X = 72
        .Pos(4).Y = 53

    End With
    
    ' 3vs3 Default
    With RetoFastConfig(3)
        .Ident = "3vs3"
        ReDim .Pos(1 To 6) As tMap
        ReDim .Users(1 To 6) As Integer
        
        .Map = 1
        .Pos(1).X = 71
        .Pos(1).Y = 44
        .Pos(2).X = 71
        .Pos(2).Y = 45
        .Pos(3).X = 71
        .Pos(3).Y = 46
        
        .Pos(4).X = 72
        .Pos(4).Y = 44
        .Pos(5).X = 72
        .Pos(5).Y = 45
        .Pos(6).X = 72
        .Pos(6).Y = 46

    End With
    
    ' # FILA DEL MEDIO
    ' # 1vs1 [ 500 ROJAS ]
    With RetoFastConfig(4)
        .Ident = "1vs1-500ROJAS"
        .Map = 1
        
        ReDim .Pos(1 To 2) As tMap
        ReDim .Users(1 To 2) As Integer
        
        .Pos(1).X = 74
        .Pos(1).Y = 49
        
        .Pos(2).X = 75
        .Pos(2).Y = 49
        
        .RedLimit = 500
        
    End With
        
    
    ' # 2vs2 [VALE RESU]
    With RetoFastConfig(5)
        .Ident = "2vs2-RESU"
        ReDim .Pos(1 To 4) As tMap
        ReDim .Users(1 To 4) As Integer
        
        .Map = 1
        .Pos(1).X = 74
        .Pos(1).Y = 52
        .Pos(2).X = 74
        .Pos(2).Y = 53
        
        .Pos(3).X = 75
        .Pos(3).Y = 52
        .Pos(4).X = 75
        .Pos(4).Y = 53
        
        .ConfigVale = ValeResu
    End With
    
    ' 3vs3 [VALE RESU]
    With RetoFastConfig(6)
        .Ident = "3vs3-RESU"
        ReDim .Pos(1 To 6) As tMap
        ReDim .Users(1 To 6) As Integer
        
        .Map = 1
        .Pos(1).X = 74
        .Pos(1).Y = 44
        .Pos(2).X = 74
        .Pos(2).Y = 45
        .Pos(3).X = 75
        .Pos(3).Y = 46
        
        .Pos(4).X = 75
        .Pos(4).Y = 44
        .Pos(5).X = 75
        .Pos(5).Y = 45
        .Pos(6).X = 75
        .Pos(6).Y = 46
        
        .ConfigVale = ValeResu
    End With
    
    ' # 3° FILA
    ' # 1vs1 [VALE TODO]
    With RetoFastConfig(7)
        .Ident = "1vs1-VALETODO"
        .Map = 1
        
        ReDim .Pos(1 To 2) As tMap
        ReDim .Users(1 To 2) As Integer
        
        .Pos(1).X = 77
        .Pos(1).Y = 49
        
        .Pos(2).X = 78
        .Pos(2).Y = 49
        
        .ConfigVale = ValeTodo
    End With
        
    
    ' # 2vs2 [VALE TODO]
    With RetoFastConfig(8)
        .Ident = "2vs2-VALETODO"
        ReDim .Pos(1 To 4) As tMap
        ReDim .Users(1 To 4) As Integer
        
        .Map = 1
        .Pos(1).X = 77
        .Pos(1).Y = 52
        .Pos(2).X = 77
        .Pos(2).Y = 53
        
        .Pos(3).X = 78
        .Pos(3).Y = 52
        .Pos(4).X = 78
        .Pos(4).Y = 53
        
        .ConfigVale = ValeTodo
    End With
    
    ' 3vs3 [VALE TODO]
    With RetoFastConfig(9)
        .Ident = "3vs3-VALETODO"
        ReDim .Pos(1 To 6) As tMap
        ReDim .Users(1 To 6) As Integer
        
        .Map = 1
        .Pos(1).X = 77
        .Pos(1).Y = 44
        .Pos(2).X = 77
        .Pos(2).Y = 45
        .Pos(3).X = 77
        .Pos(3).Y = 46
        
        .Pos(4).X = 78
        .Pos(4).Y = 44
        .Pos(5).X = 78
        .Pos(5).Y = 45
        .Pos(6).X = 78
        .Pos(6).Y = 46
        
        .ConfigVale = ValeTodo
    End With
End Sub

' # buscamos un Reto libre
Public Function RetoSlot() As Integer

    Dim A As Long
    
    For A = 1 To MAX_RETO_FAST
        With RetoFast(A)
            If .Run = False Then
                RetoSlot = A
                Exit Function
            End If
        End With
    Next A
    
End Function
' # Comprueba para enviar un nuevo reto rapido
Public Sub RetoFast_Loop()
On Error GoTo ErrHandler

    If ConfigServer.ModoRetosFast = 0 Then Exit Sub
    
    Dim A   As Long, B As Long
    Dim UserIndex() As Integer
    Dim ArenaFree As Integer
    Dim SlotReto As Integer
    Dim Check As Boolean

    For A = 1 To MAX_RETO_FAST_POSITIONS
        With RetoFastConfig(A)
            ReDim UserIndex(LBound(.Pos) To UBound(.Pos)) As Integer
                
            Check = False
            
            For B = LBound(.Users) To UBound(.Users)
            
                UserIndex(B) = MapData(.Map, .Pos(B).X, .Pos(B).Y).UserIndex
                    
                If UserIndex(B) = 0 Then Exit For
                
                If EsGm(UserIndex(B)) Then Exit For
                If UserList(UserIndex(B)).flags.Muerto Then Exit For
                If UserList(UserIndex(B)).Stats.Gld < RETOFAST_GOLD Then Exit For
                If UserList(UserIndex(B)).Stats.Points < RETOFAST_POINTS Then Exit For
                
                If B = UBound(.Pos) Then Check = True
                
                
            Next B
            
            If Check Then
                SlotReto = RetoSlot
                ArenaFree = Arenas_Free(UBound(.Pos), 0)
                If ArenaFree > 0 Then Call RetoFast_PrepareUsers(SlotReto, ArenaFree, UserIndex, .ConfigVale, .RedLimit)
            End If
            
        End With
    Next A
    
    Exit Sub

ErrHandler:
    Call LogError("Error en RetoFast_Loop")
    
End Sub

' # Comienza el Reto
Private Sub RetoFast_PrepareUsers(ByVal RetoIndex As Byte, _
                                  ByVal ArenaIndex As Integer, _
                                  ByRef UserIndex() As Integer, _
                                  ByVal ConfigVale As Byte, _
                                  ByVal RedLimit As Integer)

    On Error GoTo ErrHandler

    Dim TempOne As String

    Dim TempTwo As String

    Dim A       As Long
    
    With RetoFast(RetoIndex)
        .Run = True
        .ArenaIndex = ArenaIndex
        .Users = UserIndex
        .ConfigVale = ConfigVale
        
        .RedLimit = RedLimit
        Arenas(ArenaIndex).Used = True
        
        For A = LBound(.Users) To UBound(.Users)
            If .Users(A) > 0 Then
                With UserList(.Users(A))
                    .Stats.Gld = .Stats.Gld - RETOFAST_GOLD
                    .Stats.Points = .Stats.Points - RETOFAST_POINTS
                    .flags.SlotFast = RetoIndex
                    .flags.SlotFastUser = A
                    
                    
                    ' # Máximo de pociones que pueden usar
                    If RedLimit > 0 Then
                        .flags.RedValid = True
                        .flags.RedUsage = 0
                        .flags.RedLimit = RedLimit
                    End If
                End With
                
                WriteUpdateGold .Users(A)
                
                 If (A <= UBound(.Users) / 2) Then
                    UserList(.Users(A)).flags.FightTeam = 1
                    EventWarpUser .Users(A), Arenas(ArenaIndex).Map, Arenas(ArenaIndex).X, Arenas(ArenaIndex).Y
                Else
                    UserList(.Users(A)).flags.FightTeam = 2
                    EventWarpUser .Users(A), Arenas(ArenaIndex).Map, Arenas(ArenaIndex).X + Arenas(ArenaIndex).TileAddX, Arenas(ArenaIndex).Y + Arenas(ArenaIndex).TileAddY
                End If
                
            'Else
              'Dim NpcIndex As Integer
               ' Dim Pos As WorldPos
                'Pos.Map = Arenas(ArenaIndex).Maps(MapIndex).Map
              '  Pos.X = Arenas(ArenaIndex).Maps(MapIndex).X + Arenas(ArenaIndex).Maps(MapIndex).TileAddX
              '  Pos.Y = Arenas(ArenaIndex).Maps(MapIndex).Y + Arenas(ArenaIndex).Maps(MapIndex).TileAddY
                        
               ' NpcIndex = SpawnNpc(BOT_NPCINDEX, Pos, False, False)
               ' Npclist(NpcIndex).Inteliggence = 50 + RandomNumber(1, 50)
               ' Npclist(NpcIndex).Stats.MaxMan = 2000
               ' Npclist(NpcIndex).Stats.MaxHp = 340
               '' Npclist(NpcIndex).Stats.MinMan = Npclist(NpcIndex).Stats.MaxMan
               ' Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MaxHp
              '  Npclist(NpcIndex).BotIndex = 1
                
            End If
            
        Next A
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en RetoFast_PrepareUsers")

End Sub

' # El user abandona el reto
Public Sub RetoFast_UserAbandonate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo RetoFast_UserAbandonate_Err
        '</EhHeader>
    
        Dim SlotFast     As Byte

        Dim SlotFastUser As Byte
    
100     With UserList(UserIndex)
102         SlotFast = .flags.SlotFast
104         SlotFastUser = .flags.SlotFastUser
        
106         RetoFast(SlotFast).Users(SlotFastUser) = 0
    
        End With

        '<EhFooter>
        Exit Sub

RetoFast_UserAbandonate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetoFast.RetoFast_UserAbandonate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # El user muere en el reto
Public Sub RetoFast_UserDie(ByVal UserIndex As Integer, _
                            Optional ByVal Deslogged As Boolean = False, _
                            Optional ByVal BotIndex As Integer = 0)

    On Error GoTo ErrHandler

    Dim A            As Long

    Dim SlotFast     As Byte

    Dim SlotFastUser As Byte

    Dim TempOne      As String

    Dim TempTwo      As String
    
    Dim FightTeam As Byte
    
    With UserList(UserIndex)
        SlotFast = .flags.SlotFast
        SlotFastUser = .flags.SlotFastUser
        FightTeam = .flags.FightTeam
        
        If Deslogged Then
            Call Fast_Reset_User(SlotFast, SlotFastUser)
            RetoFast(SlotFast).Users(SlotFastUser) = 0
        End If
        
        If BotIndex > 0 Then
            Fast_Reset SlotFast
        End If
        
        For A = LBound(RetoFast(SlotFast).Users) To UBound(RetoFast(SlotFast).Users)
            If RetoFast(SlotFast).Users(A) > 0 Then
                If UserList(RetoFast(SlotFast).Users(A)).flags.FightTeam = FightTeam Then
                    If UserList(RetoFast(SlotFast).Users(A)).flags.Muerto = 0 Then Exit Sub
                End If
                
                
                If UserList(RetoFast(SlotFast).Users(A)).flags.FightTeam = 1 Then
                    TempOne = TempOne & UserList(RetoFast(SlotFast).Users(A)).Name & ", "
                Else
                    TempTwo = TempTwo & UserList(RetoFast(SlotFast).Users(A)).Name & ", "
                End If
                
                
                
                If UBound(RetoFast(SlotFast).Users) <= 2 Then
                    Call WriteUpdateEvent(UserList(RetoFast(SlotFast).Users(A)).ID, _
                            UserList(RetoFast(SlotFast).Users(A)).Name, _
                            eSubType_Modality.Fast1vs1, _
                            FightTeam <> UserList(RetoFast(SlotFast).Users(A)).flags.FightTeam)
                            
                            
                    If RetoFast(SlotFast).RedLimit > 0 Then
                        If FightTeam <> UserList(RetoFast(SlotFast).Users(A)).flags.FightTeam Then
                            Call Rachas_User_Add(RetoFast(SlotFast).Users(A))
                        Else
                            ' # Reinicia las rachas
                            UserList(RetoFast(SlotFast).Users(A)).flags.RachasTemp = 0
                        End If
                    End If
                End If
            End If
            
            
            
        Next A
        
        If Len(TempOne) Then
            TempOne = Left$(TempOne, Len(TempOne) - 2)
        Else
            TempOne = "Descalificado"
        End If
        
        If Len(TempTwo) Then
            TempTwo = Left$(TempTwo, Len(TempTwo) - 2)
        Else
            TempTwo = "Descalificado"
        End If
        
        Call Fast_Reset(SlotFast)
        
        Dim TempIdent As String
        TempIdent = (UBound(RetoFast(SlotFast).Users) / 2) & "vs" & (UBound(RetoFast(SlotFast).Users) / 2)
        

        SendData SendTarget.toMap, 1, PrepareMessageConsoleMsg( _
        "RetoFast #" & TempIdent & "» " & TempOne & " vs " & TempTwo & ". Gana " & IIf((FightTeam = 2), TempOne, TempTwo), FontTypeNames.FONTTYPE_CONSEJO)
        
        Dim TextDiscord As String
        
        
        TextDiscord = "**RetoFast #" & TempIdent & "»** " & TempOne & " vs " & TempTwo & ". **Gana " & IIf((FightTeam = 2), TempOne, TempTwo) & "**"
        
        WriteMessageDiscord CHANNEL_FIGHT, TextDiscord
    End With
    
    Exit Sub

ErrHandler:
    
    Call LogError("Error en RetoFast_UserDie")
End Sub

' # Agrega una nueva racha al personaje
Public Sub Rachas_User_Add(ByVal UserIndex As Integer)
        
    With UserList(UserIndex)
        .flags.RachasTemp = .flags.RachasTemp + 1
        
        ' # Chequea si supera las rachas historicas
        If .flags.RachasTemp > .flags.Rachas Then
            .flags.Rachas = .flags.RachasTemp
        End If
        
        ' # Mensaje SPAM de rachas temporales cada 20 rondas
        If .flags.RachasTemp Mod 20 = 0 Then
            Call SendData(SendTarget.ToAll, 0, _
                PrepareMessageConsoleMsg("Rachas» El personaje '" & UCase$(.Name) & "' alcanza las " & .flags.RachasTemp & " rachas. (Record: " & .flags.Rachas & ")", FontTypeNames.FONTTYPE_USERPREMIUM))
        End If
        
        
        ' # Mensaje SPAM de rachas historicas cada 100
        If .flags.Rachas Mod 100 = 0 Then
            Call SendData(SendTarget.ToAll, 0, _
                PrepareMessageConsoleMsg("Rachas» '" & UCase$(.Name) & "' alcanza las " & .flags.Rachas & " rachas.", FontTypeNames.FONTTYPE_ANGEL))
        End If
        
        
        ' # Sonido de Rachas
        Call Rachas_Add_Sound(UserIndex)
        
    End With
    
    ' # Guardamos el Personaje en la DB
    Call WriteUpdateUserData(UserList(UserIndex))
End Sub

Public Sub Rachas_Add_Sound(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim Sound As Integer
    
    With UserList(UserIndex)
    
        ' # Sonidos PREMIUM por tener RECORD DE RACHAS 50+
        If .flags.Rachas > 50 Then
            Select Case .flags.RachasTemp
                Case 1: Sound = 262
                Case 2: Sound = 261
                Case 3: Sound = 263
                Case 5: Sound = 273
                Case 10: Sound = 271
            End Select
        End If
        
        
        ' # Sonidos

        Select Case .flags.RachasTemp
            Case 25: Sound = 260
            Case 50: Sound = 266
            Case 75: Sound = 269
            Case 100: Sound = 267
            Case 125: Sound = 270
            Case 150: Sound = 264
            Case 175: Sound = 265
            Case 200: Sound = 268
            Case 250: Sound = 272
        End Select
        
    End With
    
    If Sound > 0 Then
        Call SendData(SendTarget.toMap, 1, PrepareMessagePlayEffect(Sound, NO_3D_SOUND, NO_3D_SOUND))
    End If
    
ErrHandler:
End Sub

' # Reiniciamos variables del User
Public Sub Fast_Reset_User(ByVal Slot As Byte, ByVal SlotUser As Byte)
        '<EhHeader>
        On Error GoTo Fast_Reset_User_Err
        '</EhHeader>
    
100     With RetoFast(Slot)
102         UserList(.Users(SlotUser)).flags.SlotFast = 0
104         UserList(.Users(SlotUser)).flags.SlotFastUser = 0
106         UserList(.Users(SlotUser)).flags.FightTeam = 0
            
            UserList(.Users(SlotUser)).flags.RedValid = False
            UserList(.Users(SlotUser)).flags.RedUsage = 0
            UserList(.Users(SlotUser)).flags.RedLimit = 0
        
108         EventWarpUser .Users(SlotUser), 1, 73 + SlotUser, 73
        End With
                
        '<EhFooter>
        Exit Sub

Fast_Reset_User_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetoFast.Fast_Reset_User " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Reiniciamos un Reto especifico
Public Sub Fast_Reset(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo Fast_Reset_Err
        '</EhHeader>

        Dim A As Long
    
100     With RetoFast(Slot)
            If .ArenaIndex > 0 Then
102         For A = LBound(.Users) To UBound(.Users)

104             If .Users(A) > 0 Then
            
106                 Call Fast_Reset_User(Slot, A)
108                 .Users(A) = 0
                
                End If

110         Next A
        
112         .Run = False
            Arenas(.ArenaIndex).Used = False
            .ArenaIndex = 0
            End If
        End With

        '<EhFooter>
        Exit Sub

Fast_Reset_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetoFast.Fast_Reset " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Reiniciamos todos los retos rapidos vigentes
Public Sub Fast_Reset_All()
        '<EhHeader>
        On Error GoTo Fast_Reset_All_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_RETO_FAST
102         Call Fast_Reset(A)
104     Next A
    
        '<EhFooter>
        Exit Sub

Fast_Reset_All_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRetoFast.Fast_Reset_All " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


