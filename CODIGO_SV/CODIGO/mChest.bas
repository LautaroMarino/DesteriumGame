Attribute VB_Name = "mChest"
' // Todo este procesamiento debe ir en un servidor externo el día de mañana.

Option Explicit

Public Type tChestData
    Map As Integer
    X As Byte
    Y As Byte
    
    ObjIndex As Integer ' Objeto que saldrá de la tierra
    Time As Long
End Type

Public Const MAX_CHESTDATA As Integer = 500

Public ChestLast As Integer
Public ChestData(1 To MAX_CHESTDATA) As tChestData

' Busca un Slot libre para agregar le cofre al conteo de respawn
Private Function ChestData_Slot() As Integer
        '<EhHeader>
        On Error GoTo ChestData_Slot_Err
        '</EhHeader>
        Dim A As Long
    
100     For A = 1 To MAX_CHESTDATA
102         If ChestData(A).Map = 0 Then
104             ChestData_Slot = A
                Exit Function
        
            End If
106     Next A
        '<EhFooter>
        Exit Function

ChestData_Slot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mChest.ChestData_Slot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function ChestData_Add(ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte, _
                         ByVal ObjIndex As Integer, _
                         ByVal Time As Long, _
                         ByVal DropObj As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo ChestData_Add_Err
        '</EhHeader>
    
        Dim Slot As Byte
    
100     Slot = ChestData_Slot
    
102     If Slot = 0 Then
104         Call LogError("¡¡ERROR AL AGREGAR UN COFRE EN EL MAPA " & Map & " " & X & " " & Y)
        Else
        
        
106         If Not DropObj Then
108             Time = Time * 1.5 ' 50% más de tiempo para que regenere en caso de haber roto
            End If
        
110         With ChestData(Slot)
112             .Map = Map
114             .X = X
116             .Y = Y
118             .Time = Time
120             .ObjIndex = ObjIndex
            End With
        
122         ChestData_Add = True
        End If

        '<EhFooter>
        Exit Function

ChestData_Add_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mChest.ChestData_Add " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


Public Sub ChestLoop()
        '<EhHeader>
        On Error GoTo ChestLoop_Err
        '</EhHeader>
        Dim A As Long
        Dim ChestNull As tChestData
        Dim Obj As Obj
    
100     For A = 1 To MAX_CHESTDATA
102         With ChestData(A)
            
104             If .Map > 0 Then
106                 .Time = .Time - 1

108                 If .Time = 0 Then
110                     Obj.ObjIndex = .ObjIndex
112                     Obj.Amount = 1
                    
114                     Call EraseObj(MapData(.Map, .X, .Y).ObjInfo.Amount, .Map, .X, .Y)
116                     Call MakeObj(Obj, .Map, .X, .Y)
118                     Call SendToAreaByPos(.Map, .X, .Y, PrepareMessagePlayEffect(eSound.sChestClose, .X, .Y))
120                     ChestData(A) = ChestNull
                
                    End If
            
                End If
        
            End With
    
    
122     Next A

        '<EhFooter>
        Exit Sub

ChestLoop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mChest.ChestLoop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Chest_DropObj(ByVal UserIndex As Integer, _
                         ByVal ObjIndex As Integer, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte, _
                         ByVal DropInv As Boolean)

        '<EhHeader>
        On Error GoTo Chest_DropObj_Err

        '</EhHeader>
    
        Dim DropObj    As Obj

        Dim A          As Long

        Dim RandomDrop As Byte

        Dim nPos       As WorldPos

        Dim Pos        As WorldPos
    
        Dim Sound      As Integer
    
        Dim Random     As Byte
            
100     Random = RandomNumber(1, ObjData(ObjIndex).Chest.NroDrop)
            
        RandomDrop = ObjData(ObjIndex).Chest.Drop(Random)
102     Pos.Map = Map
104     Pos.X = X
106     Pos.Y = Y
    
108     With DropData(RandomDrop)

110         For A = 1 To .Last
112             DropObj.ObjIndex = .Data(A).ObjIndex
114             DropObj.Amount = RandomNumber(.Data(A).Amount(0), .Data(A).Amount(1))

                If RandomNumber(1, 100) <= .Data(A).Prob Then
                    If DropInv Then
115                     If Not MeterItemEnInventario(UserIndex, DropObj) Then
                            Call TirarItemAlPiso(UserList(UserIndex).Pos, DropObj)
    
                        End If
    
                    Else
                        Call Tilelibre(Pos, nPos, DropObj, False, True)
    
118                     If nPos.X <> 0 And nPos.Y <> 0 Then
120                         Call MakeObj(DropObj, nPos.Map, nPos.X, nPos.Y)
122                         nPos.Map = 0
124                         nPos.X = 0
126                         nPos.Y = 0
    
                        End If
    
                    End If

                End If
                
128         Next A
            
            Call Chest_PlaySound(UserIndex, X, Y)

        End With
    
        '<EhFooter>
        Exit Sub

Chest_DropObj_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.Chest_DropObj " & "at line " & Erl

        '</EhFooter>
End Sub

Public Sub Chest_AbreFortuna(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo Chest_AbreFortuna_Err
        '</EhHeader>
        Dim Num As Long
        Dim Obj As Obj

100     Num = RandomNumber(1, ObjData(ObjIndex).MaxFortunas)

102     If Not MeterItemEnInventario(UserIndex, ObjData(ObjIndex).Fortuna(Num)) Then
104         Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjData(ObjIndex).Fortuna(Num))

        End If
                    
106     Call WriteConsoleMsg(UserIndex, "¡Has recibido " & ObjData(ObjData(ObjIndex).Fortuna(Num).ObjIndex).Name & " (x" & ObjData(ObjIndex).Fortuna(Num).Amount & ")!", FontTypeNames.FONTTYPE_INFOGREEN)
108     Call Chest_PlaySound(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
        '<EhFooter>
        Exit Sub

Chest_AbreFortuna_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mChest.Chest_AbreFortuna " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Chest_PlaySound(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo Chest_PlaySound_Err
        '</EhHeader>

        Dim Random As Byte
        Dim Sound As Long
    
100     Random = RandomNumber(1, 100)
        
102     If Random <= 25 Then
104         Sound = eSound.sChestDrop1
106     ElseIf Random <= 50 Then
108         Sound = eSound.sChestDrop2
        Else
110         Sound = eSound.sChestDrop3

        End If
        
112     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Sound, X, Y))

        '<EhFooter>
        Exit Sub

Chest_PlaySound_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mChest.Chest_PlaySound " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
