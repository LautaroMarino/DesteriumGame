Attribute VB_Name = "EventosDS_Reward"
' Los usuarios podrán compartir sus objetos y donarlos para los eventos predeterminados del juego

Option Explicit

' Buscamos un Slot Libre para agregar el objeto
Private Function Events_Reward_Slot(ByVal SlotEvent As Byte) As Byte
        '<EhHeader>
        On Error GoTo Events_Reward_Slot_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_REWARD_OBJ
102         If Events(SlotEvent).RewardObj(A).ObjIndex = 0 Then
104             Events_Reward_Slot = A
                Exit For
            End If
106     Next A
        '<EhFooter>
        Exit Function

Events_Reward_Slot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.EventosDS_Reward.Events_Reward_Slot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
' Agregamos un nuevo objeto a la lista de PREMIOS DONADOS
Public Sub Events_Reward_Add(ByVal UserIndex As Integer, _
                             ByVal SlotEvent As Byte, _
                             ByVal Slot As Byte, _
                             ByVal Amount As Integer)
        '<EhHeader>
        On Error GoTo Events_Reward_Add_Err
        '</EhHeader>
                             
        Dim ObjIndex   As Integer
        Dim A          As Long
        Dim SlotReward As Byte
    
        ' Chequeamos que haya lugar para guardar el nuevo objeto
100     If Events(SlotEvent).LastReward = MAX_REWARD_OBJ Then
102         Call WriteConsoleMsg(UserIndex, "No hay más espacio para agregar premios al evento.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
104     ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
106     SlotReward = Events_Reward_Slot(SlotEvent)
    
        ' Quitamos el Objeto al Usuario
108     Call QuitarUserInvItem(UserIndex, Slot, Amount)
110     Call UpdateUserInv(False, UserIndex, Slot)
        
        ' Agregamos el Objeto a la lista de premios
112     Events(SlotEvent).RewardObj(SlotReward).ObjIndex = ObjIndex
114     Events(SlotEvent).RewardObj(SlotReward).Amount = Amount
        Events(SlotEvent).LastReward = Events(SlotEvent).LastReward + 1
116     Call WriteConsoleMsg(UserIndex, "Has donado para el evento " & Events(SlotEvent).Name & " el objeto " & ObjData(ObjIndex).Name & " (x" & Amount & ")", FontTypeNames.FONTTYPE_INFOGREEN)
    
118     LogEventos "El personaje " & UserList(UserIndex).Name & " Ha donado para el evento " & Events(SlotEvent).Name & " el objeto " & ObjData(ObjIndex).Name & " (x" & Amount & ")"
        '<EhFooter>
        Exit Sub

Events_Reward_Add_Err:
        LogError Err.description & vbCrLf & _
               "in Events_Reward_Add " & _
               "at line " & Erl

        '</EhFooter>
End Sub

' Recorre la lista de premios donados y en caso de caer sobre uno válido, se lo "regala" al personaje campeón.
Public Sub Events_Reward_User(ByVal UserIndex As Integer, ByVal SlotEvent As Byte)
        '<EhHeader>
        On Error GoTo Events_Reward_User_Err
        '</EhHeader>
    
        Dim A As Long
        Dim Slot As Byte
    
100     With Events(SlotEvent)
102         Slot = RandomNumber(1, MAX_REWARD_OBJ)
        
104         If .RewardObj(Slot).ObjIndex > 0 Then
106             If Not MeterItemEnInventario(UserIndex, .RewardObj(Slot)) Then
108                 WriteConsoleMsg UserIndex, "Tu premio Donado no ha sido entregado, envia esta foto a un Game Master.", FontTypeNames.FONTTYPE_INFO
110                 LogEventos ("Personaje " & UserList(UserIndex).Name & " no recibió: " & .RewardObj(Slot).ObjIndex & " (x" & .RewardObj(Slot).Amount & ")")
                    Exit Sub
                End If
            
112             .RewardObj(Slot).ObjIndex = 0
114             .RewardObj(Slot).Amount = 0
            
116             Call WriteConsoleMsg(UserIndex, "Un premio donado ha caído sobre tí por haber ganado el evento ¡Esto no siempre sucede, Felicitaciones!", FontTypeNames.FONTTYPE_INFOGREEN)
        
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

Events_Reward_User_Err:
        LogError Err.description & vbCrLf & _
               "in Events_Reward_User " & _
               "at line " & Erl

        '</EhFooter>
End Sub

