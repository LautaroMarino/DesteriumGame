Attribute VB_Name = "mReward"
Option Explicit

Public Type tPlayerData
    PlayerName As String
    GamesWon As Integer
    GamesPlayed As Integer
    ConsecutiveWins As Integer
End Type

' # Selecciona el Premio a dar según el evento. Se podrian poner otros parámetros.
Public Function Reward_Detect(ByVal Modality As eSubType_Modality, _
                        ByRef Player As tPlayerData, _
                        ByVal Clase As Byte) As Obj
    
    Dim IObj As Obj
    IObj.Amount = 1         ' # Por el momento recibe una armadura
    
    Select Case Modality
        Case eSubType_Modality.General
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.unoVSuno
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.dosVSdos
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.tresVStres
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.DagaRusa
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.DeathMatch
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.Imparable
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.ReyVsRey
            IObj.ObjIndex = 0
            
        Case eSubType_Modality.Retos1vs1
            If Clase = eClass.Paladin Or Clase = eClass.Cleric Or Clase = eClass.Warrior Then
                IObj.ObjIndex = 2250
            Else
                IObj.ObjIndex = 2101
            End If
            
        Case eSubType_Modality.Fast1vs1
            If Clase = eClass.Paladin Or Clase = eClass.Cleric Or Clase = eClass.Warrior Then
                IObj.ObjIndex = 2249
            Else
                IObj.ObjIndex = 2100
            End If
    End Select
    
    Reward_Detect = IObj
End Function

' # Descuento la recompensa y se la quita al final
Public Sub Reward_Check_User(ByVal UserIndex As Integer)
    Dim A As Long
    
    Static Seconds As Long
    
    With UserList(UserIndex)
       For A = 1 To .Stats.BonusLast
            With .Stats.Bonus(A)
                If .DurationSeconds > 0 Then
                    .DurationSeconds = .DurationSeconds - 1
                    
                    If .DurationSeconds = 0 Then
                        Call Reward_User_Remove(UserIndex, A)
                    End If
                End If
                    
                If .DurationDate <> vbNullString Then
                    If Seconds >= 30 Then
                        Seconds = 0
                        If DateDiff("s", Now, .DurationDate) <= 0 Then
106                         Call Reward_User_Remove(UserIndex, A)
                        End If
                    End If
                End If
                
            End With
       Next A
       
       Seconds = Seconds + 1
        
        
       If Seconds > 10000 Then Seconds = 0
    End With
End Sub

' # Quita la bonificación del usuario
Private Sub Reward_User_Remove(ByVal UserIndex As Integer, ByVal Slot As Integer)

    With UserList(UserIndex)
  
        With .Stats.Bonus(Slot)
            If .Tipo = eObj Then
                Call WriteConsoleMsg(UserIndex, "¡El Item ha desaparecido de tu inventario! Ha pasado el tiempo de duración.", FontTypeNames.FONTTYPE_INFORED)
                Call QuitarObjetos(.Value, .Amount, UserIndex)
            End If
        
            .Amount = 0
            .DurationDate = 0
            .DurationSeconds = 0
            .Tipo = 0
            .Value = 0
        End With
        
        'If .Stats.BonusLast = Slot Then
            '.Stats.BonusLast = .Stats.BonusLast - 1
            'ReDim Preserve .Stats.Bonus(1 To .Stats.BonusLast) As UserBonus
        'End If
        

    End With
End Sub
' # Procesa la recompensa del usuario
Public Sub Reward_Process_User(ByVal ModalityID As Byte, _
                               ByRef Player As tPlayerData)
    
    On Error GoTo ErrHandler
    
    Dim tUser As Integer
    Dim IObj As Obj
    Dim FilePath As String
    Dim TempDate As String
    Dim Clase As Byte
    
    tUser = NameIndex(Player.PlayerName)
    
    FilePath = CharPath & UCase$(Player.PlayerName) & ".chr"
    If tUser > 0 Then
        Clase = UserList(tUser).Clase
    Else
        Clase = val(GetVar(FilePath, "INIT", "CLASE"))
    End If
    
    If Clase = 0 Then Exit Sub
    
     ' # Detecta que item dar en cada caso.
    IObj = Reward_Detect(ModalityID, Player, Clase)
    
    If IObj.ObjIndex = 0 Then Exit Sub ' # No tiene items para dar
    
    ' # 1 mes de duración
    TempDate = Detect_FirstDayNext
    
    If tUser > 0 Then
        Call Bonus_AddUser_Online(tUser, eBonusType.eObj, IObj.ObjIndex, IObj.Amount, 0, TempDate)
        Call Log_Reward("Procesando online a " & Player.PlayerName & " en " & ModalityID)
    Else
        Call Bonus_AddUser_Offline(Player.PlayerName, eBonusType.eObj, IObj.ObjIndex, IObj.Amount, 0, TempDate)
        Call Log_Reward("Procesando offline a " & Player.PlayerName & " en " & ModalityID)
    End If
  
    Call Log_Reward(Player.PlayerName & " recibió el objeto " & ObjData(IObj.ObjIndex).Name & " (x" & IObj.Amount & ")")
    
    Exit Sub
ErrHandler:
    
End Sub

' # Busca un Slot vacio en el inventario del usuario
Private Function Reward_InventorySlot_Offline(ByRef Manager As clsIniManager, ByVal ObjIndex As Integer)
    Dim A As Long
    Dim Temp As String
    Dim Amount As Integer
    
    ' # Busca un Slot repetido
    For A = 1 To MAX_INVENTORY_SLOTS
        Temp = Manager.GetValue("INVENTORY", "OBJ" & A)
        
        If val(ReadField(1, Temp, Asc("-"))) = ObjIndex Then
        
            ' # Se duda que supere los 10.000 pero por si acaso.
            Amount = val(ReadField(2, Temp, Asc("-")))
            
            If Amount < MAX_INVENTORY_OBJS Then
            
                Reward_InventorySlot_Offline = A
                Exit Function
            End If
        End If
    Next A
    
    ' # Busca un slot Libre
    For A = 1 To MAX_INVENTORY_SLOTS
        Temp = Manager.GetValue("INVENTORY", "OBJ" & A)
        
        If val(ReadField(1, Temp, Asc("-"))) = 0 Then
            Reward_InventorySlot_Offline = A
            Exit Function
        End If
    Next A
End Function

' # Busca un Slot Vacio para asignar el BONUS [ONLINE]
Public Function Bonus_User_SearchSlot(ByVal UserIndex As Integer) As Integer

    Dim A As Long

    With UserList(UserIndex).Stats
        For A = 1 To .BonusLast
            If .Bonus(A).Tipo = 0 Then
                Bonus_User_SearchSlot = A
                Exit Function
            End If
        Next A
    End With
End Function

' # Busca un Slot Vacio para asignar el BONUS [OFFLINE]
Public Function Bonus_User_SearchSlot_Offline(ByVal BonusLast As Integer, ByRef Manager As clsIniManager) As Integer
    Dim A As Long
    Dim Temp As String

    For A = 1 To BonusLast
        Temp = Manager.GetValue("BONUS", "BONUS" & A)
        
        If val(ReadField(1, Temp, Asc("|"))) = 0 Then
            Bonus_User_SearchSlot_Offline = A
            Exit Function
        End If
    Next A
End Function
' # Agrega un Bonus nuevo al personaje
Public Sub Bonus_AddUser_Online(ByVal UserIndex As Integer, _
                              ByRef Tipo As eBonusType, _
                              ByVal Value As Long, _
                              ByVal Amount As Long, _
                              ByVal DurationSeconds As Long, _
                              ByVal DurationDate As String, _
                              Optional ByVal EntregaObjeto As Boolean = True)
                              
On Error GoTo ErrHandler

                             
    Dim IObj As Obj
    Dim Stats As UserStats
    Dim Bonus As tBonus
    Dim Slot As Integer
    
    With UserList(UserIndex)
        If .Invent.NroItems = MAX_INVENTORY_SLOTS Then
            Call Log_Reward("El inventario de " & .Name & " está ocupado. Objeto NO entregado.")
            Exit Sub
        End If
10
        Slot = Bonus_User_SearchSlot(UserIndex)
20
        If Slot = 0 Then
            .Stats.BonusLast = .Stats.BonusLast + 1
            
            ReDim Preserve .Stats.Bonus(0 To .Stats.BonusLast) As UserBonus
            Slot = .Stats.BonusLast
        End If
30
        With .Stats.Bonus(Slot)
            .Tipo = Tipo
            .Value = Value
            .Amount = Amount
            .DurationDate = DurationDate
            .DurationSeconds = DurationSeconds
        End With
40
        Bonus.Tipo = Tipo
        Bonus.Porc = Value
50
        If Bonus.Tipo = eBonusType.eObj Then
            IObj.ObjIndex = Value
            IObj.Amount = Amount
            
            If EntregaObjeto Then
                If Not MeterItemEnInventario(UserIndex, IObj) Then
                    Call Log_Reward("El inventario de " & .Name & " está ocupado. Objeto NO entregado.")
                    Exit Sub
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Has recibido " & ObjData(IObj.ObjIndex).Name & " en tu inventario. Duración: " & DurationDate, FontTypeNames.FONTTYPE_INVASION)
        End If
60
    End With

    Exit Sub
ErrHandler:
    Call LogError("Bonus_Add error en linea " & Erl)
End Sub

' # Agrega un Bonus nuevo al personaje (Offline)
Public Sub Bonus_AddUser_Offline(ByVal Name As String, _
                              ByRef Tipo As eBonusType, _
                              ByVal Value As Long, _
                              ByVal Amount As Long, _
                              ByVal DurationSeconds As Long, _
                              ByVal DurationDate As String)
                              
                              
    On Error GoTo ErrHandler
    
    Dim Stats As UserStats
    Dim Bonus As tBonus
    Dim tUser As Integer
    Dim Slot As Integer
    Dim SlotInvent As Integer
    
    Dim Manager As clsIniManager
    Set Manager = New clsIniManager
    
    Dim FilePath As String
    FilePath = CharPath & UCase$(Name) & ".chr"
10
    If Not FileExist(FilePath) Then Exit Sub
20
    Manager.Initialize FilePath
30
    With Stats
        .BonusLast = val(Manager.GetValue("BONUS", "BONUSLAST"))
40
        SlotInvent = Reward_InventorySlot_Offline(Manager, Value)
50
        If SlotInvent = 0 Then
            Call Log_Reward("El inventario de " & Name & " está ocupado. Objeto NO entregado.")
            Exit Sub
        End If
60
        Slot = Bonus_User_SearchSlot_Offline(.BonusLast, Manager)
70
        If Slot = 0 Then
        
            Slot = .BonusLast + 1
            Call Manager.ChangeValue("BONUS", "BONUSLAST", CStr(Slot))
        End If
80
        
        Dim SlotsOcupados As Integer
        Dim AmountSlot As Integer
        
        ' Objeto nuevo +1 cantidad items
        If Manager.GetValue("INVENTORY", "OBJ" & SlotInvent) = "0-0-0" Then
            SlotsOcupados = val(Manager.GetValue("INVENTORY", "CANTIDADITEMS"))
            
            Call Manager.ChangeValue("INVENTORY", "CANTIDADITEMS", CStr(SlotsOcupados + 1))
        End If
        
        Call Manager.ChangeValue("INVENTORY", "OBJ" & SlotInvent, Value & "-" & Amount & "-0")
90
        Call Manager.ChangeValue("BONUS", "BONUS" & Slot, CStr(Tipo) & "|" & CStr(Value) & "|" & CStr(Amount) & "|" & _
                                                          CStr(DurationSeconds) & "|" & CStr(DurationDate))
    End With
    
100
    Manager.DumpFile FilePath
    
    Set Manager = New clsIniManager
    
    Exit Sub
ErrHandler:
    Call LogError("Bonus_AddUser_Offline:: Error al procesar offline bonus in " & Erl)
End Sub

