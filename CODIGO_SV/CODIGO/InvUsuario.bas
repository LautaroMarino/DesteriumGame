Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    '17/09/02
    'Agregue que la función se asegure que el objeto no es un barco

    On Error GoTo ErrHandler

    Dim i        As Integer

    Dim ObjIndex As Integer
    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
            ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And _
            ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And _
            ObjData(ObjIndex).Bronce <> 1 And Not ItemNewbie(ObjIndex)) Then
                TieneObjetosRobables = True

                Exit Function

            End If
        End If

    Next i
    
    Exit Function

ErrHandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.number & " - " & Err.description)
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo manejador
    
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then

            Dim i As Integer

            For i = 1 To NUMCLASES

                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                    ClasePuedeUsarItem = False
                    sMotivo = "Tu clase no puede usar este objeto."

                    Exit Function

                End If

            Next i

        End If
    End If
    
    ClasePuedeUsarItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function
Function ClasePuedeItem(ByVal Clase As Integer, _
                        ByVal ObjIndex As Integer) As Boolean
    On Error GoTo manejador
    
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then

            Dim i As Integer

            For i = 1 To NUMCLASES

                If ObjData(ObjIndex).ClaseProhibida(i) = Clase Then
                    ClasePuedeItem = False
                    Exit Function

                End If

            Next i

        End If
    
    ClasePuedeItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeItem")
End Function

' Comprueba si tiene objetos que para su level no está permitido usar más...
Sub QuitarLevelObj(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo QuitarLevelObj_Err
        '</EhHeader>
    
        Dim j As Long
    
100     With UserList(UserIndex)
102         For j = 1 To .CurrentInventorySlots

104             If .Invent.Object(j).ObjIndex > 0 Then
106                 If ObjData(.Invent.Object(j).ObjIndex).LvlMax > 0 Then
108                     If .Stats.Elv >= ObjData(.Invent.Object(j).ObjIndex).LvlMax Then
110                         Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
112                         Call UpdateUserInv(False, UserIndex, j)
                        End If
                    End If
                End If

114         Next j
    
        End With

        '<EhFooter>
        Exit Sub

QuitarLevelObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.QuitarLevelObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Sub QuitarNewbieObj(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuitarNewbieObj_Err
        '</EhHeader>

        Dim j As Integer

100     With UserList(UserIndex)

102         For j = 1 To UserList(UserIndex).CurrentInventorySlots

104             If .Invent.Object(j).ObjIndex > 0 Then
106                 If ObjData(.Invent.Object(j).ObjIndex).Newbie = 1 Then
108                     Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
110                     Call UpdateUserInv(False, UserIndex, j)
                    End If
                End If

112         Next j

114         'If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
        
116             'Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
    
            'End If
        End With

        '<EhFooter>
        Exit Sub

QuitarNewbieObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.QuitarNewbieObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LimpiarInventario_Err
        '</EhHeader>

        Dim j As Integer

100     With UserList(UserIndex)

102         For j = 1 To .CurrentInventorySlots
104             .Invent.Object(j).ObjIndex = 0
106             .Invent.Object(j).Amount = 0
108             .Invent.Object(j).Equipped = 0
110         Next j
    
112         .Invent.NroItems = 0
    
114         .Invent.ArmourEqpObjIndex = 0
116         .Invent.ArmourEqpSlot = 0
    
118         .Invent.WeaponEqpObjIndex = 0
120         .Invent.WeaponEqpSlot = 0
    
122         .Invent.AuraEqpObjIndex = 0
124         .Invent.AuraEqpSlot = 0
    
126         .Invent.CascoEqpObjIndex = 0
128         .Invent.CascoEqpSlot = 0
    
130         .Invent.EscudoEqpObjIndex = 0
132         .Invent.EscudoEqpSlot = 0
    
134         .Invent.AnilloEqpObjIndex = 0
136         .Invent.AnilloEqpSlot = 0
    
138         .Invent.MunicionEqpObjIndex = 0
140         .Invent.MunicionEqpSlot = 0
    
142         .Invent.BarcoObjIndex = 0
144         .Invent.BarcoSlot = 0
    
146         .Invent.MochilaEqpObjIndex = 0
148         .Invent.MochilaEqpSlot = 0
    
150         .Invent.MonturaObjIndex = 0
152         .Invent.MochilaEqpSlot = 0
    
154         .Invent.MagicObjIndex = 0
156         .Invent.MagicSlot = 0
        End With

        '<EhFooter>
        Exit Sub

LimpiarInventario_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.LimpiarInventario " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim A As Long
    
    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
    
        ' En eventos de cambio de clase,raza,level los objetos no se consumen. Excepto las Pociones
        A = UserList(UserIndex).flags.SlotEvent

        If A > 0 Then
            If Events(A).ChangeClass > 0 Or Events(A).ChangeRaze > 0 Or Events(A).ChangeLevel > 0 Then
                If (ObjData(.ObjIndex).OBJType = otFlechas) Then Exit Sub
            End If
        End If

        If .Amount <= cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)

        End If
        
        'Quita un objeto
        .Amount = .Amount - cantidad

        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            
            .ObjIndex = 0
            .Amount = 0

        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                  ByVal UserIndex As Integer, _
                  ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim NullObj As UserOBJ

    Dim LoopC   As Long

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then
    
            'Actualiza el inventario
            If .Invent.Object(Slot).ObjIndex > 0 Then
                Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
            Else
                Call ChangeUserInv(UserIndex, Slot, NullObj)
            End If
    
        Else
    
            'Actualiza todos los slots
            For LoopC = 1 To .CurrentInventorySlots

                'Actualiza el inventario
                If .Invent.Object(LoopC).ObjIndex > 0 Then
                    Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
                Else
                    Call ChangeUserInv(UserIndex, LoopC, NullObj)
                End If
            
            Next LoopC

        End If
    
        Exit Sub

    End With

ErrHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.number & " : " & Err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal Slot As Byte, _
            ByVal Num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: 11/5/2010
        '11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
        '***************************************************
        '<EhHeader>
        On Error GoTo DropObj_Err

        '</EhHeader>

        Dim DropObj  As Obj

        Dim MapObj   As Obj

        Dim TempTick As Long
        
100     With UserList(UserIndex)
            TempTick = GetTime

102         If Num > 0 Then
        
104             DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex

106             If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otMonturas Then
108                 Call WriteConsoleMsg(UserIndex, "No puedes tirar la montura.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
110             If Not EsGmDios(UserIndex) Then
112                 If ObjData(DropObj.ObjIndex).NoNada = 1 Then
                        If ObjData(DropObj.ObjIndex).LvlMax > UserList(UserIndex).Stats.Elv Then
                            Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
                            Call UpdateUserInv(False, UserIndex, Slot)
                        Else
                             Call WriteConsoleMsg(UserIndex, "No puedes realizar ninguna acción con este objeto. ¡Podría ser de uso personal!", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        Exit Sub
    
                    End If

                End If
            
116             If ObjData(DropObj.ObjIndex).NoDrop = 1 Then
118                 Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
120             If Not EsGm(UserIndex) Then
122                 If ObjData(DropObj.ObjIndex).Premium = 1 Then
124                     Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los objetos Premium!!", FontTypeNames.FONTTYPE_TALK)
                    
                        Exit Sub

                    End If
            
126                 If ObjData(DropObj.ObjIndex).Oro = 1 Then
128                     Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los objetos Oro!!", FontTypeNames.FONTTYPE_TALK)
                    
                        Exit Sub

                    End If
                
130                 If ObjData(DropObj.ObjIndex).Plata = 1 Then
132                     Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los objetos Plata!!", FontTypeNames.FONTTYPE_TALK)
                    
                        Exit Sub

                    End If
                
134                 If ObjData(DropObj.ObjIndex).OBJType = otTransformVIP Then
136                     Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los skins!!", FontTypeNames.FONTTYPE_TALK)

                        Exit Sub

                    End If
                
138                 If ObjData(DropObj.ObjIndex).Caos = 1 Or ObjData(DropObj.ObjIndex).Real = 1 Then
140                     Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos faccionarios.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                End If
            
                ' // NUEVO
142             If .flags.SlotReto > 0 Then
144                 If Retos(.flags.SlotReto).config(eRetoConfig.eItems) = 1 Then
146                     Call WriteConsoleMsg(UserIndex, "No puedes dropear objetos si estas luchando por los mismos", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                End If
            
152             If .flags.SlotEvent > 0 Then
154                 If Events(.flags.SlotEvent).LimitRed > 0 Then
156                     Call WriteConsoleMsg(UserIndex, "No puedes dropear objetos si estas luchando por límite de pociones rojas.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                End If
                
                If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otBarcos Then
                     Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡NO puedes tirar los barcos al suelo!", FontTypeNames.FONTTYPE_TALK)
                    'Exit Sub
                End If
                    
            
158             DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)

                'Check objeto en el suelo
160             MapObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
162             MapObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
        
164             If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
        
166                 If MapObj.Amount = MAX_INVENTORY_OBJS Then
168                     Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
            
170                 If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
172                     DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount

                    End If
            
174                 If Not ItemNewbie(DropObj.ObjIndex) Then Call MakeObj(DropObj, Map, X, Y)
176                 Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
178                 Call UpdateUserInv(False, UserIndex, Slot)
            
180
            
184                 If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otGemas Then
                        If TempTick - .Counters.SpamMessage > 60000 Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje '" & .Name & "' ha tirado la " & ObjData(DropObj.ObjIndex).Name & " en " & MapInfo(.Pos.Map).Name & "(Mapa: " & .Pos.Map & " " & .Pos.X & " " & .Pos.Y & ")", FontTypeNames.FONTTYPE_GUILD))
                            .Counters.SpamMessage = TempTick

                        End If

186

                    End If
            
188                 If Not .flags.Privilegios And PlayerType.User Then
190                     Call Logs_User(.Name, eGm, eDropObj, "tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                    End If

192                 If ObjData(DropObj.ObjIndex).Log = 1 Then
194                     Call Logs_User(.Name, eLog.eUser, eDropObj, "tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                
196                 ElseIf DropObj.Amount > 100 Then

198                     If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
200                         Call Logs_User(.Name, eLog.eUser, eDropObj, "tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                        End If

                    End If

                Else
202                 Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With

        '<EhFooter>
        Exit Sub

DropObj_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.DropObj " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub EraseObj(ByVal Num As Integer, _
             ByVal Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EraseObj_Err
        '</EhHeader>

100     With MapData(Map, X, Y)
102         .ObjInfo.Amount = .ObjInfo.Amount - Num
    
104         If .ObjInfo.Amount <= 0 Then
106             .ObjInfo.ObjIndex = 0
108             .ObjInfo.Amount = 0
110             .ObjEvent = 0
            
112             Call ModAreas.DeleteEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT)
            End If

        End With

        '<EhFooter>
        Exit Sub

EraseObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.EraseObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub MakeObj(ByRef Obj As Obj, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo MakeObj_Err
        '</EhHeader>
    
100     If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then
    
102         With MapData(Map, X, Y)
            
104             If .ObjInfo.ObjIndex = Obj.ObjIndex Then
106                 .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
                Else
108                 .Protect = GetTime
110                 .ObjInfo = Obj
                
112                 If .trigger <> eTrigger.zonaOscura Then
                         Dim Coordinates As WorldPos
114                     Coordinates.Map = Map
116                     Coordinates.X = X
118                     Coordinates.Y = Y
                
120                     Call ModAreas.CreateEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT, Coordinates, ObjData(.ObjInfo.ObjIndex).SizeWidth, ObjData(.ObjInfo.ObjIndex).SizeHeight)
                    End If
                End If

            End With

        End If

        '<EhFooter>
        Exit Sub

MakeObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.MakeObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj, Optional ByVal ShowMessage As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim Slot As Byte
    
    With UserList(UserIndex)
        '¿el user ya tiene un objeto del mismo tipo?
        Slot = 1
        
        Do Until .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            If Slot > .CurrentInventorySlots Then
                Exit Do
            End If
        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    If ShowMessage Then Call WriteConsoleMsg(UserIndex, "No puedes cargar más objetos.", FontTypeNames.FONTTYPE_FIGHT)
                    MeterItemEnInventario = False
                    Exit Function
                End If

            Loop

            .Invent.NroItems = .Invent.NroItems + 1
        End If
    
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot <= MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.ObjIndex) Then
                If ShowMessage Then Call WriteConsoleMsg(UserIndex, "No puedes contener objetos especiales en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                MeterItemEnInventario = False
                Exit Function
            End If
        End If

        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
            .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
        End If

    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Function

ErrHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.number & " : " & Err.description)
End Function

Sub GetObj(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 18/12/2009
        '18/12/2009: ZaMa - Oro directo a la billetera.
        '***************************************************
        '<EhHeader>
        On Error GoTo GetObj_Err
        '</EhHeader>

        Dim Obj    As ObjData

        Dim MiObj  As Obj

        Dim ObjPos As String

100     With UserList(UserIndex)

            '¿Hay algun obj?
102         If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then
104           If Not EsGm(UserIndex) Then
106                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1 Then Exit Sub
                End If
            
                '¿Esta permitido agarrar este obj?
108             If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

                    Dim X As Integer

                    Dim Y As Integer
                
110                 X = .Pos.X
112                 Y = .Pos.Y
                
114                 Obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
116                 MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
118                 MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
120                 MapData(.Pos.Map, .Pos.X, .Pos.Y).Protect = 0
                
                    ' Oro directo a la billetera!
                    'If Obj.OBJType = otGuita Then
                    ' .Stats.Gld = .Stats.Gld + MiObj.Amount
                    'Quitamos el objeto
                    'Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                        
                    ' Call WriteUpdateGold(UserIndex)
                    'Else
122                 If MeterItemEnInventario(UserIndex, MiObj) Then
                         
                        'Quitamos el objeto
124                     Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)

                          ' Comprobamos si esta en una misión
                          Call Quests_Check_Objs(UserIndex, MiObj.ObjIndex, MiObj.Amount)
                          
126                     If Not .flags.Privilegios And PlayerType.User Then
128                         Call Logs_User(.Name, eGm, eGetObj, .Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                        End If
        
130                     If ObjData(MiObj.ObjIndex).Log = 1 Then
132                         ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
134                         Call Logs_User(.Name, eLog.eUser, eGetObj, .Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                            
136                     ElseIf MiObj.Amount > 100 Then

138                         If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
140                             ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
142                             Call Logs_User(.Name, eLog.eUser, eGetObj, .Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                            End If
                        End If
                    End If

                    'End If
                End If

            Else
144             Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

GetObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.GetObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: 26/05/2011
    '26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
    '***************************************************

    On Error GoTo ErrHandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(UserIndex)
        With .Invent

            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then

                Exit Sub

            ElseIf .Object(Slot).ObjIndex = 0 Then

                Exit Sub

            End If
            
            Obj = ObjData(.Object(Slot).ObjIndex)
        End With
        
        If Obj.SkillNum > 0 Or Obj.SkillsEspecialNum > 0 Then
            Call UserStats_UpdateEffectAll(UserIndex, Obj, False)
        End If
        
        Select Case Obj.OBJType
            
            Case eOBJType.otMonturas
                Call DoEquita(UserIndex, Obj, Slot)
                
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MonturaObjIndex = 0
                    .MonturaSlot = 0
                End With
                
            Case eOBJType.otReliquias
                ' mEffect.Effect_UpdateUser UserIndex, True
                
                With .Invent
                    .Object(Slot).Equipped = 0
                    .ReliquiaObjIndex = 0
                    .ReliquiaSlot = 0
                End With
                
            Case eOBJType.otPendienteParty
                ' # Actualizar los porcentajes
                If .GroupIndex > 0 Then
                    UpdatePorcentaje .GroupIndex
                End If
                
                With .Invent
                    .Object(Slot).Equipped = 0
                    .PendientePartyObjIndex = 0
                    .PendientePartySlot = 0
                End With
                
            Case eOBJType.otMagic

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MagicObjIndex = 0
                    .MagicSlot = 0
                End With
                
            Case eOBJType.otWeapon

                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0
                End With
                
                
                 '.Skins.WeaponIndex = 0
                 
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .AuraIndex(2) = 0
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)
                    End With

                End If
                
            Case eOBJType.otAuras

                With .Invent
                    .Object(Slot).Equipped = 0
                    .AuraEqpObjIndex = 0
                    .AuraEqpSlot = 0
                End With
            
                With .Char
                    .AuraIndex(5) = 0
                    Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)
                End With
                
            Case eOBJType.otFlechas

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
            
            Case eOBJType.otAnillo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0
                End With
            
            Case eOBJType.otarmadura
                
                
                If .flags.TransformVIP > 0 Then
                    Call TransformVIP_User(UserIndex, 0)
                End If
                With .Invent

                    'Si tiene armadura faccionaria de segunda jerarquía equipada la sacamos:
                    If .FactionArmourEqpObjIndex Then
                        Call Desequipar(UserIndex, .FactionArmourEqpSlot)
                    End If
                    
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0
                    
                    
                End With
                
                '.Skins.ArmourIndex = 0
                
                If .flags.Navegando = 0 Then
                    Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                End If
                
                With .Char
                    .AuraIndex(1) = 0
                    Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)
                End With
                 
            Case eOBJType.otcasco

                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0
                End With
                
                ' .Skins.HelmIndex = 0
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .AuraIndex(3) = 0
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)
                    End With

                End If
            
            Case eOBJType.otescudo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0
                End With
                
                ' .Skins.ShieldIndex = 0
                 
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                         .AuraIndex(4) = 0
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)
                    End With

                End If
            
            Case eOBJType.otMochilas

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0
                End With
                
                Call InvUsuario.TirarTodosLosItemsEnMochila(UserIndex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End Select

    End With
    
    Call WriteUpdateUserStats(UserIndex)
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Desquipar. Error " & Err.number & " : " & Err.description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).Mujer = 1 Then
            SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
        ElseIf ObjData(ObjIndex).Hombre = 1 Then
            SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
        Else
            SexoPuedeUsarItem = True
        End If
        
    Else
        SexoPuedeUsarItem = True
    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."

    Exit Function

ErrHandler:
    Call LogError("SexoPuedeUsarItem")
End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer, _
                              Optional ByRef sMotivo As String) As Boolean
        '<EhHeader>
        On Error GoTo FaccionPuedeUsarItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 26/05/2011 (Amraphen)
        '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
        '26/05/2011: Amraphen - Agrego validación para armaduras faccionarias de segunda jerarquía.
        '***************************************************
        Dim ArmourIndex           As Integer

        Dim FaltaPrimeraJerarquia As Boolean

100     If ObjData(ObjIndex).Real Then
102         If Not Escriminal(UserIndex) And esArmada(UserIndex) Then
104             If ObjData(ObjIndex).Real = 2 Then
106                 ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
108                 If ArmourIndex > 0 And ObjData(ArmourIndex).Real = 1 Then
110                     FaccionPuedeUsarItem = True
                    Else
112                     FaccionPuedeUsarItem = False
114                     FaltaPrimeraJerarquia = True
                    End If

                Else 'Es item faccionario común
116                 FaccionPuedeUsarItem = True
                End If

            Else
118             FaccionPuedeUsarItem = False
            End If

120     ElseIf ObjData(ObjIndex).Caos Then

122         If Escriminal(UserIndex) And esCaos(UserIndex) Then
124             If ObjData(ObjIndex).Caos = 2 Then
126                 ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
128                 If ArmourIndex > 0 And ObjData(ArmourIndex).Caos = 1 Then
130                     FaccionPuedeUsarItem = True
                    Else
132                     FaccionPuedeUsarItem = False
134                     FaltaPrimeraJerarquia = True
                    End If

                Else 'Es item faccionario común
136                 FaccionPuedeUsarItem = True
                End If

            Else
138             FaccionPuedeUsarItem = False
            End If

        Else
140         FaccionPuedeUsarItem = True
        End If
    
142     If Not FaccionPuedeUsarItem Then
144         If FaltaPrimeraJerarquia Then
146             sMotivo = "Debes tener equipada una armadura faccionaria."
            Else
148             sMotivo = "Tu alinación no puede usar este objeto."
            End If
        End If

        '<EhFooter>
        Exit Function

FaccionPuedeUsarItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.FaccionPuedeUsarItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function CheckUserSkill(ByVal UserIndex As Integer, _
                                ByRef Obj As ObjData) As Boolean

        '<EhHeader>
        On Error GoTo CheckUserSkill_Err

        '</EhHeader>

100     With UserList(UserIndex)

104         If .Stats.UserSkills(eSkill.Magia) < Obj.MagiaSkill Then
106             Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.MagiaSkill & " skills en Mágia.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If

110         If .Stats.UserSkills(eSkill.Resistencia) < Obj.RMSkill Then
112             Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.RMSkill & " skills en Resistencia Mágica.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If

114         If Obj.OBJType = otWeapon Then
116             If .Stats.UserSkills(eSkill.Armas) < Obj.ArmaSkill Then
118                 Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaSkill & " skills en Combate con Armas.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If

            End If

122         If .Stats.UserSkills(eSkill.Defensa) < Obj.EscudoSkill Then
124             Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.EscudoSkill & " skills en Defensa con Escudos.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If

128         If .Stats.UserSkills(eSkill.Tacticas) < Obj.ArmaduraSkill Then
130             Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaduraSkill & " skills en Evasión.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If

132         If Obj.OBJType = otWeapon Then
134             If .Stats.UserSkills(eSkill.Proyectiles) < Obj.ArcoSkill Then
136                 Call WriteConsoleMsg(UserIndex, "Para usar este item tienes que tener " & Obj.ArcoSkill & " skills en Armas de Proyectiles.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If

            End If

140         If .Stats.UserSkills(eSkill.Apuñalar) < Obj.DagaSkill Then
142             Call WriteConsoleMsg(UserIndex, "Para utilizar este ítem necesitas " & Obj.DagaSkill & " skills en Apuñalar.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If
            
            If .Stats.UserSkills(eSkill.Magia) < Obj.MagiaSkill Then
                Call WriteConsoleMsg(UserIndex, "Para usar este item tienes que tener " & Obj.MagiaSkill & " skills en Magia.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If
        
144         CheckUserSkill = True
    
        End With

        '<EhFooter>
        Exit Function

CheckUserSkill_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.CheckUserSkill " & "at line " & Erl

        

        '</EhFooter>
End Function

' @ Aplicamos /Des aplicamops los atributos del objeto
Sub UserStats_UpdateEffectAll(ByVal UserIndex As Integer, _
                              ByRef Obj As ObjData, _
                              ByVal Equipped As Boolean)
        '<EhHeader>
        On Error GoTo UserStats_UpdateEffectAll_Err
        '</EhHeader>
    
        Dim A          As Long

        Dim SkillIndex As Integer

        Dim Amount     As Integer
    
100     With Obj
    
102         If .SkillNum > 0 Then

104             For A = 1 To .SkillNum
106                 SkillIndex = .Skill(A).Selected
108                 Amount = IIf(Equipped, .Skill(A).Amount, -.Skill(A).Amount)
110                 UserList(UserIndex).Stats.UserSkills(.Skill(A).Selected) = UserList(UserIndex).Stats.UserSkills(.Skill(A).Selected) + Amount
112             Next A

            End If
        
114         If .SkillsEspecialNum > 0 Then

116             For A = 1 To .SkillsEspecialNum
118                 SkillIndex = .SkillsEspecial(A).Selected
120                 Amount = IIf(Equipped, .SkillsEspecial(A).Amount, -.SkillsEspecial(A).Amount)
122                 UserList(UserIndex).Stats.UserSkillsEspecial(SkillIndex) = UserList(UserIndex).Stats.UserSkillsEspecial(SkillIndex) + Amount
                       Call UserStats_UpdateEffectUser(UserIndex, SkillIndex, Amount)
124             Next A
        
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

UserStats_UpdateEffectAll_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.UserStats_UpdateEffectAll " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub UserStats_UpdateEffectUser(ByVal UserIndex As Integer, _
                               ByVal SkillNum As Byte, _
                               ByVal Amount As Integer)

        '<EhHeader>
        On Error GoTo UserStats_UpdateEffectUser_Err

        '</EhHeader>
    
100     With UserList(UserIndex)

102         Select Case SkillNum
        
                Case 1 ' Vida
                    .Stats.MaxHp = .Stats.MaxHp + Amount
                    .Stats.MinHp = .Stats.MaxHp
                    Call WriteUpdateUserStats(UserIndex)
                    
104             Case 2 ' Maná
                    .Stats.MaxMan = .Stats.MaxMan + Amount
                    .Stats.MinMan = .Stats.MaxMan
                    Call WriteUpdateUserStats(UserIndex)
                    
106             Case 3 'Curación : Skill que define un porcentaje (1 a 100)
            
108             Case 4 'Escudo Mágico : Skill que define un porcentaje (1 a 100)
            
110             Case 5 'Veneno : Skill que define un porcentaje (1 a 100)
            
112             Case 6 'Fuerza
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + Amount
                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        Call WriteUpdateStrenght(UserIndex)
                        
114             Case 7 'Agilidad
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + Amount
                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Fuerza)
214                   Call WriteUpdateDexterity(UserIndex)
            End Select
    
        End With

        '<EhFooter>
        Exit Sub

UserStats_UpdateEffectUser_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.UserStats_UpdateEffectUser " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '*************************************************
    'Author: Unknown
    'Last modified: 26/05/2011 (Amraphen)
    '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
    '14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
    '26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
    '*************************************************

    On Error GoTo ErrHandler

    'Equipa un item del inventario
    Dim Obj      As ObjData

    Dim ObjIndex As Integer

    Dim sMotivo  As String
    
    With UserList(UserIndex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
        Obj = ObjData(ObjIndex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
                
        If Obj.Bronce = 1 And Not .flags.Bronce = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [AVENTURERO] pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Plata = 1 And Not .flags.Plata = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [HEROE] pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Oro = 1 And Not .flags.Oro = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [LEYENDA] pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Premium = 1 And Not .flags.Premium = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios PREMIUM pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Navidad = 1 And ModoNavidad = 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya supera navidad wei", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
                
        If Obj.LvlMax <> 0 And Obj.LvlMax < .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto hasta Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
            Exit Sub

        End If
        
        If Obj.LvlMin <> 0 And Obj.LvlMin > .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto a partir del Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
            Exit Sub

        End If
        
        ' Skill requerido para el objeto
        If Not CheckUserSkill(UserIndex, Obj) Then Exit Sub
        
        If .flags.SlotReto > 0 Then
        
            ' Uso de Escudos/Cascos
            If (Retos(.flags.SlotReto).config(eRetoConfig.eEscudos) = 0 And Obj.OBJType = otescudo) Or (Retos(.flags.SlotReto).config(eRetoConfig.eCascos) = 0 And Obj.OBJType = otcasco) Then
                Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el uso de este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            ' Uso de Objetos [BRONCE] [PLATA] [ORO] [PREMIUM]
            'If (Retos(.flags.SlotReto).config(eRetoConfig.eBronce) = 0 And .flags.Bronce = 0) Or (Retos(.flags.SlotReto).config(eRetoConfig.ePlata) = 0 And .flags.Plata = 0) Or (Retos(.flags.SlotReto).config(eRetoConfig.eOro) = 0 And .flags.Oro = 0) Or (Retos(.flags.SlotReto).config(eRetoConfig.ePremium) = 0 And .flags.Premium = 0) Then
                
                'Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el uso de este objeto.", FontTypeNames.FONTTYPE_INFO)

               ' Exit Sub

           ' End If
            
        End If
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DagaRusa Then
                Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el uso de objetos.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        End If

        Select Case Obj.OBJType

            Case eOBJType.otMagic

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.MagicObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MagicSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MagicObjIndex = ObjIndex
                    .Invent.MagicSlot = Slot
                    
                Else
                    
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If

            Case eOBJType.otReliquias

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.ReliquiaObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ReliquiaSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ReliquiaObjIndex = ObjIndex
                    .Invent.ReliquiaSlot = Slot
                    
                    'Call mEffect.Effect_UpdateUser(UserIndex, False)
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eOBJType.otPendienteParty
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.PendientePartyObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.PendientePartySlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.PendientePartyObjIndex = ObjIndex
                    .Invent.PendientePartySlot = Slot
                    
                    Call WriteConsoleMsg(UserIndex, "En caso de que seas líder de un grupo podrás cambiar el porcentaje hasta " & ObjData(ObjIndex).Porc & "%", FontTypeNames.FONTTYPE_INFOGREEN)
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eOBJType.otAuras

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                     
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.AuraIndex(5) = NingunAura
                        Else
                            .Char.AuraIndex(5) = NingunAura
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub

                    End If
                     
                    'Quitamos el elemento anterior
                    If .Invent.AuraEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.AuraEqpSlot)

                    End If
                     
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AuraEqpObjIndex = ObjIndex
                    .Invent.AuraEqpSlot = Slot
                     
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = ObjData(ObjIndex).AuraIndex
                        .CharMimetizado.AuraIndex(5) = ObjData(ObjIndex).AuraIndex
                    Else
                        .Char.AuraIndex(5) = ObjData(ObjIndex).AuraIndex
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eOBJType.otWeapon

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub
                    Else

                        ' Quiere equipar un arma dos manos y tiene escudo.
                        If .Invent.EscudoEqpObjIndex > 0 Then
                            If ObjData(ObjIndex).DosManos = 1 Then
                                Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                            End If

                        End If

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = ObjIndex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_SACARARMA, .Pos.X, .Pos.Y, .Char.charindex))
                    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, ObjIndex)
                        .CharMimetizado.AuraIndex(2) = ObjData(ObjIndex).AuraIndex(2)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, ObjIndex)
                        .Char.AuraIndex(2) = ObjData(ObjIndex).AuraIndex(2)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
                'Call Skins_CheckObj(UserIndex, ObjIndex)

            Case eOBJType.otAnillo

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.AnilloEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

                    End If
                
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otFlechas

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                        
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

                    End If
                
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MunicionEqpObjIndex = ObjIndex
                    .Invent.MunicionEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otarmadura

                If .flags.Navegando = 1 Then Exit Sub
                If .flags.Montando = 1 Then Exit Sub
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And SexoPuedeUsarItem(UserIndex, ObjIndex, sMotivo) And CheckRazaUsaRopa(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Nos fijamos si es armadura de segunda jerarquia
                    If Obj.Real = 2 Or Obj.Caos = 2 Then

                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            Call Desequipar(UserIndex, Slot)
                            
                            If Not .flags.Mimetizado = 1 Then
                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                            End If
                            
                            Exit Sub

                        End If
                        
                        'Quita el anterior
                        If .Invent.FactionArmourEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.FactionArmourEqpSlot)

                        End If
                        
                        'Lo equipa
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.FactionArmourEqpObjIndex = ObjIndex
                        .Invent.FactionArmourEqpSlot = Slot
                        
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.Body = GetArmourAnim(UserIndex, ObjIndex)
                            
                        End If

                    Else

                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            Call Desequipar(UserIndex, Slot)
                            
                            'Esto está de más:
                            'Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                            If Not .flags.Mimetizado = 1 Then
                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                            End If
                            
                            Exit Sub

                        End If
                
                        'Quita el anterior
                        If .Invent.ArmourEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                        End If
                
                        'Lo equipa
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.ArmourEqpObjIndex = ObjIndex
                        .Invent.ArmourEqpSlot = Slot
                            
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.Body = GetArmourAnim(UserIndex, ObjIndex)
                            .CharMimetizado.AuraIndex(1) = ObjData(ObjIndex).AuraIndex(1)
                        Else
                            .Char.Body = GetArmourAnim(UserIndex, ObjIndex)
                            .Char.AuraIndex(1) = ObjData(ObjIndex).AuraIndex(1)
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        .flags.Desnudo = 0

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
                'Call Skins_CheckObj(UserIndex, ObjIndex)
                
            Case eOBJType.otcasco

                If .flags.Navegando = 1 Then Exit Sub
                
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).config(eConfigEvent.eCascoEscudo) = 0 Then Exit Sub

                End If
        
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub

                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = ObjIndex
                    .Invent.CascoEqpSlot = Slot

                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = GetHelmAnim(UserIndex, .Invent.CascoEqpObjIndex)
                        .CharMimetizado.AuraIndex(3) = ObjData(ObjIndex).AuraIndex(3)
                    Else
                        .Char.CascoAnim = GetHelmAnim(UserIndex, .Invent.CascoEqpObjIndex)
                        .Char.AuraIndex(3) = ObjData(ObjIndex).AuraIndex(3)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
                'Call Skins_CheckObj(UserIndex, ObjIndex)

            Case eOBJType.otescudo

                If .flags.Navegando = 1 Then Exit Sub
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).config(eConfigEvent.eCascoEscudo) = 0 Then Exit Sub

                End If
                
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
        
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.ShieldAnim = NingunEscudo
                        Else
                            .Char.ShieldAnim = NingunEscudo
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub
                        
                    Else

                        ' Quiere equipar un escudo y tiene arma dos manos
                        If .Invent.WeaponEqpObjIndex > 0 Then
                            If ObjData(.Invent.WeaponEqpObjIndex).DosManos = 1 Then
                                Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                            End If

                        End If

                    End If
             
                    'Quita el anterior
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                    End If
             
                    'Lo equipa
                     
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.EscudoEqpObjIndex = ObjIndex
                    .Invent.EscudoEqpSlot = Slot
                     
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.ShieldAnim = GetShieldAnim(UserIndex, .Invent.EscudoEqpObjIndex)
                        .CharMimetizado.AuraIndex(4) = ObjData(ObjIndex).AuraIndex(4)
                    Else
                        .Char.ShieldAnim = GetShieldAnim(UserIndex, ObjIndex)
                        .Char.AuraIndex(4) = ObjData(ObjIndex).AuraIndex(4)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                 
                'Call Skins_CheckObj(UserIndex, ObjIndex)
    
            Case eOBJType.otMochilas

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)

                    Exit Sub

                End If

                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.MochilaEqpSlot)

                End If

                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = ObjIndex
                .Invent.MochilaEqpSlot = Slot
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                Call WriteAddSlots(UserIndex, Obj.MochilaType)

        End Select
    
    End With

    ' Agrega los Atributos necesarios segun los skills del objeto.
    If Obj.SkillNum > 0 Or Obj.SkillsEspecialNum > 0 Then
        Call UserStats_UpdateEffectAll(UserIndex, Obj, True)

    End If
        
    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub
    
ErrHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.number & " - Error Description : " & Err.description)

End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, _
                                 ItemIndex As Integer, _
                                 Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        
        
        'Verifica si la raza puede usar la ropa
        If .Raza = eRaza.Humano Or .Raza = eRaza.Elfo Or .Raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
            If (ObjData(ItemIndex).RopajeEnano <> 0) Then
                CheckRazaUsaRopa = True
            Else
                CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
            End If
        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.Raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False
        End If

    End With
    
    If EsGm(UserIndex) Then CheckRazaUsaRopa = True
    
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
    Exit Function
    
ErrHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function


Private Sub Potion_SimulatePotion(ByVal UserIndex As Integer, _
                                  ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo Potion_SimulatePotion_Err
        '</EhHeader>
                                  
                                  
        Dim TempTick As Long
    
100     With UserList(UserIndex)
            'Quitamos del inv el item
102         Call QuitarUserInvItem(UserIndex, Slot, 1)
104         Call UpdateUserInv(False, UserIndex, Slot)

            ' Los admin invisibles solo producen sonidos a si mismos
106         If .flags.AdminInvisible = 1 Then
108             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
            Else
110             If TempTick - .Counters.RuidoPocion > 1000 Then
112                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
114                 .Counters.RuidoPocion = TempTick
                End If
            End If
        End With
        '<EhFooter>
        Exit Sub

Potion_SimulatePotion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.Potion_SimulatePotion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UseInvItem(ByVal UserIndex As Integer, _
               ByVal Slot As Byte, _
               ByVal SecondaryClick As Byte, _
               ByVal Value As Long)

        '*************************************************
        'Author: Unknown
        'Last modified: 10/12/2009
        'Handels the usage of items from inventory box.
        '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
        '24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
        '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
        '17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
        '27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
        '08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
        '10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
        '*************************************************
        '<EhHeader>
        On Error GoTo UseInvItem_Err

        '</EhHeader>

        Dim Obj      As ObjData

        Dim ObjIndex As Integer

        Dim TargObj  As ObjData

        Dim MiObj    As Obj

        Dim sMotivo  As String
    
100     With UserList(UserIndex)
    
102         If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
104         Obj = ObjData(.Invent.Object(Slot).ObjIndex)

106         If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
108             Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If
        
110         If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex, sMotivo) Then
112             Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If
        
114         If Obj.OBJType = otTransformVIP Then
116             If Not CheckRazaUsaRopa(UserIndex, .Invent.Object(Slot).ObjIndex, sMotivo) Then
118                 Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFORED)
    
                    Exit Sub
    
                End If

            End If
        
120         If Obj.Bronce = 1 And Not .flags.Bronce = 1 Then
122             Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [AVENTURERO] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERBRONCE)

                Exit Sub

            End If
        
124         If Obj.Plata = 1 And Not .flags.Plata = 1 Then
126             Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [HEROE] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERPLATA)

                Exit Sub

            End If
        
128         If Obj.Oro = 1 And Not .flags.Oro = 1 Then
130             Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [LEYENDA] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERGOLD)

                Exit Sub

            End If
        
132         If Obj.Premium = 1 And Not .flags.Premium = 1 Then
134             Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [PREMIUM] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERPREMIUM)

                Exit Sub

            End If
        
136         If Obj.LvlMax <> 0 And Obj.LvlMax < .Stats.Elv Then
138             Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto hasta Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
                Exit Sub

            End If
        
            If Obj.LvlMin <> 0 And Obj.LvlMin > .Stats.Elv Then
                Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto a partir del Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
                Exit Sub

            End If
            
140         If Obj.OBJType = eOBJType.otWeapon Then
142             If Obj.proyectil = 1 Then
                
                    'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
144                 If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
                
                    If Obj.Municion = 1 Then
146                     If .Invent.MunicionEqpObjIndex = 0 Then
148                         Call WriteConsoleMsg(UserIndex, "Debes equipar las municiones antes de usar el arma de proyectil.", FontTypeNames.FONTTYPE_USERPREMIUM)

                            Exit Sub

                        End If

                    End If

                Else

                    'dagas
150                 If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

                End If

            Else

152             If SecondaryClick Then
154                 If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
                Else

156                 If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

                End If
           
            End If
        
            If .flags.Meditando Then
                .flags.Meditando = False
                .Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))

            End If
            
158         ObjIndex = .Invent.Object(Slot).ObjIndex
160         .flags.TargetObjInvIndex = ObjIndex
162         .flags.TargetObjInvSlot = Slot
        
164         Select Case Obj.OBJType
                
                Case eOBJType.otItemRandom
                    
                    Call Chest_AbreFortuna(UserIndex, ObjIndex)
                    
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                    Call UpdateUserInv(False, UserIndex, Slot)
                    
                Case eOBJType.otcofre
                
                    Call Chest_DropObj(UserIndex, ObjIndex, 0, 0, 0, True)
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                    Call UpdateUserInv(False, UserIndex, Slot)

                Case eOBJType.otPociones

166                 If .flags.Muerto = 1 Then
168                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
170                 If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then Exit Sub
                
                    Dim TempTick As Long, CanUse As Boolean

172                 .flags.TomoPocion = True
174                 .flags.TipoPocion = Obj.TipoPocion
                
176                 TempTick = GetTime
178                 CanUse = True
                
180                 Select Case .flags.TipoPocion
                
                        Case 1 'Modif la agilidad
182                         .flags.DuracionEfecto = Obj.DuracionEfecto
                
                            'Usa el item
184                         .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

186                         If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS

188                         If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        
                            'Quitamos del inv el item
190                         If Obj.Ilimitado = 0 Then
192                             Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
194                         If .flags.AdminInvisible = 1 Then
196                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            
                            Else

198                             If TempTick - .Counters.RuidoPocion > 1000 Then
200                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
202                                 .Counters.RuidoPocion = TempTick

                                End If

                            End If

204                         Call WriteUpdateDexterity(UserIndex)
                        
206                     Case 2 'Modif la fuerza
208                         .flags.DuracionEfecto = Obj.DuracionEfecto
                
                            'Usa el item
210                         .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

212                         If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS

214                         If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        
                            'Quitamos del inv el item
216                         If Obj.Ilimitado = 0 Then
218                             Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
220                         If .flags.AdminInvisible = 1 Then
222                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            Else

224                             If TempTick - .Counters.RuidoPocion > 1000 Then
226                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
228                                 .Counters.RuidoPocion = TempTick

                                End If

                            End If

230                         Call WriteUpdateStrenght(UserIndex)
                        
232                     Case 3 'Pocion roja, restaura HP
                            
                            ' # Está en un evento que cuenta las rojas
                            If .flags.RedValid Then
                                .flags.RedUsage = .flags.RedUsage + 1
                                
                                If .flags.RedUsage > .flags.RedLimit Then
                                    Call WriteConsoleMsg(UserIndex, "Parece ser que el evento tiene limite de pociones rojas a configurado a un máximo de: " & .flags.RedLimit, FontTypeNames.FONTTYPE_INFORED)
                                    Exit Sub
                                End If
                            End If
                            
                            
234                         If CanUse Then
236                             .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
    
238                             If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp

                            End If
                        
                            'Quitamos del inv el item
240                         If Obj.Ilimitado = 0 Then
242                             Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
244                         If .flags.AdminInvisible = 1 Then
246                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            Else

248                             If TempTick - .Counters.RuidoPocion > 1000 Then
250                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
252                                 .Counters.RuidoPocion = TempTick

                                End If

                            End If
                        
254                         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
                        
256                         Call WriteUpdateHP(UserIndex)

                            
                            
258                     Case 4 'Pocion azul, restaura MANA
                        
260                         If CanUse Then
262                             .Stats.MinMan = .Stats.MinMan + (Porcentaje(.Stats.MaxMan, 3) + .Stats.Elv \ 2 + 40 / .Stats.Elv)
                            
264                             If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan

                            End If
                        
                            'Quitamos del inv el item
266                         If Obj.Ilimitado = 0 Then
268                             Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
270                         If .flags.AdminInvisible = 1 Then
272                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            Else

274                             If TempTick - .Counters.RuidoPocion > 1000 Then
276                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
278                                 .Counters.RuidoPocion = TempTick

                                End If

                            End If
                        
280                         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
                        
282                         Call WriteUpdateMana(UserIndex)

284                     Case 5 ' Pocion violeta

286                         If .flags.Envenenado = 1 Then
288                             .flags.Envenenado = 0
290                             Call WriteUpdateEffect(UserIndex)
292                             Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)

                            End If

                            'Quitamos del inv el item
294                         If Obj.Ilimitado = 0 Then
296                             Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
298                         If .flags.AdminInvisible = 1 Then
300                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            Else

302                             If TempTick - .Counters.RuidoPocion > 1000 Then
304                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
306                                 .Counters.RuidoPocion = TempTick

                                End If

                            End If
                        
308                         Call WriteUpdateUserStats(UserIndex)
                        
310                     Case 6  ' Pocion Negra

312                         If .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Sub
314                         If .flags.Comerciando Then Exit Sub
                        
316                         If .flags.Privilegios And PlayerType.User Then
318                             Call QuitarUserInvItem(UserIndex, Slot, 1)
320                             Call UserDie(UserIndex)
322                             Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)

                            End If
                        
324                         Call WriteUpdateUserStats(UserIndex)
                        
326                     Case 7 ' Poción de energía

328                         If .flags.Transform = 1 Then
330                             Call WriteConsoleMsg(UserIndex, "No puedes utilizar esta poción estando transformado.", FontTypeNames.FONTTYPE_INFORED)

                                Exit Sub

                            End If
                        
332                         If Obj.Ilimitado = 0 Then
334                             Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If
                              
                            ' Los admin invisibles solo producen sonidos a si mismos
336                         If .flags.AdminInvisible = 1 Then
338                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            Else

340                             If TempTick - .Counters.RuidoPocion > 1000 Then
342                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
344                                 .Counters.RuidoPocion = TempTick

                                End If
        
                            End If
                              
346                         .Stats.MinSta = .Stats.MinSta + (.Stats.MaxSta * 0.1)

348                         If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                              
350                         Call WriteUpdateSta(UserIndex)

                    End Select
               
352                 Call UpdateUserInv(False, UserIndex, Slot)
                
354             Case eOBJType.oteffect
356                 Call mEffect.Effect_Add(UserIndex, Slot, .Invent.Object(Slot).ObjIndex)
            
358             Case eOBJType.otUseOnce

360                 If .flags.Muerto = 1 Then
362                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
        
                    'Usa el item
364                 .Stats.MinHam = .Stats.MinHam + Obj.MinHam

366                 If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
368                 .flags.Hambre = 0
370                 Call WriteUpdateHungerAndThirst(UserIndex)
                    'Sonido
                
372                 If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
374                     Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
                    Else
376                     Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)

                    End If
                
                    'Quitamos del inv el item
378                 Call QuitarUserInvItem(UserIndex, Slot, 1)
                
380                 Call UpdateUserInv(False, UserIndex, Slot)
        
382             Case eOBJType.otGuita

384                 If .flags.Muerto = 1 Then
386                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
388                 .Stats.Gld = .Stats.Gld + .Invent.Object(Slot).Amount
390                 .Invent.Object(Slot).Amount = 0
392                 .Invent.Object(Slot).ObjIndex = 0
394                 .Invent.NroItems = .Invent.NroItems - 1
                
396                 Call UpdateUserInv(False, UserIndex, Slot)
398                 Call WriteUpdateGold(UserIndex)
                
400             Case eOBJType.otGuitaDsp

402                 If .flags.Muerto = 1 Then
404                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
406                 .Stats.Eldhir = .Stats.Eldhir + .Invent.Object(Slot).Amount
408                 .Invent.Object(Slot).Amount = 0
410                 .Invent.Object(Slot).ObjIndex = 0
412                 .Invent.NroItems = .Invent.NroItems - 1
                
414                 Call UpdateUserInv(False, UserIndex, Slot)
416                 Call WriteUpdateDsp(UserIndex)

418             Case eOBJType.otGemasEffect

420                 If .flags.Muerto Then
422                     Call WriteConsoleMsg(UserIndex, "No puedes usar bonificaciones estando muerto.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
424                 If .flags.SelectedBono > 0 Then
426                     WriteConsoleMsg UserIndex, "Ya tienes un efecto activado.", FontTypeNames.FONTTYPE_INFO

                        Exit Sub

                    End If
                
428                 .flags.SelectedBono = .Invent.Object(Slot).ObjIndex
430                 .Counters.TimeBono = ObjData(.Invent.Object(Slot).ObjIndex).BonoTime * 60
                
432                 WriteConsoleMsg UserIndex, "Has activado el efecto de la gema. El mismo desaparecerá en " & Int(.Counters.TimeBono / 60) & " minutos. Utiliza /EST para saber cuando tiempo te queda.", FontTypeNames.FONTTYPE_INFO

434             Case eOBJType.otGemaTelep

436                 If Obj.TelepMap = 0 Or Obj.TelepX = 0 Or Obj.TelepY = 0 Then Exit Sub
438                 If .flags.Muerto Then Exit Sub
440                 If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Or .flags.Desafiando > 0 Or .Counters.Pena > 0 Then Exit Sub
442                 If MapInfo(.Pos.Map).Pk Then Exit Sub
                
444                 If .flags.Plata = 0 Then
446                     Call WriteConsoleMsg(UserIndex, "Necesitas ser usuario Plata para utilizar este scroll.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                
448                 If .flags.ObjIndex > 0 Then
450                     If .flags.ObjIndex <> .Invent.Object(Slot).ObjIndex Then
452                         Call WriteConsoleMsg(UserIndex, "Ya tienes activado otro efecto. Haz clic sobre el objeto correspondiente", FontTypeNames.FONTTYPE_INFORED)

                            Exit Sub

                        End If
                    
454                     Call WriteConsoleMsg(UserIndex, "Has regresado al mapa.", FontTypeNames.FONTTYPE_INFOGREEN)
456                     WarpUserChar UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, False
                    Else
458                     .flags.ObjIndex = .Invent.Object(Slot).ObjIndex
                    
460                     If .flags.Premium > 0 Then
462                         Obj.TelepTime = Obj.TelepTime + 10

                        End If
                    
464                     .Counters.TimeTelep = Obj.TelepTime * 60
466                     WarpUserChar UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, False
468                     WriteConsoleMsg UserIndex, "Has activado el efecto de la teletransportación.", FontTypeNames.FONTTYPE_INFO

                    End If
                        
                    'Quitamos del inv el item
                    'Call QuitarUserInvItem(UserIndex, Slot, 1)
                    'Call UpdateUserInv(False, UserIndex, Slot)

470             Case eOBJType.otWeapon

472                 If .flags.Muerto = 1 Then
474                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    If Not .Stats.MinSta > 5 And .Counters.Trabajando > 0 Then
                        .Counters.Trabajando = 0
                            
                        Call WriteUpdateUserTrabajo(UserIndex)

                    End If
                        
476                 If Not .Stats.MinSta > 0 Then
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
480                 If ObjData(ObjIndex).proyectil = 1 Then
482                     If .Invent.Object(Slot).Equipped = 0 Then
484                         Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If

486                     Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)

                    Else
                    
488                     Select Case ObjIndex
                    
                            Case CAÑA_PESCA, RED_PESCA, CAÑA_COFRES
                            
                                ' Lo tiene equipado?
490                             If Not .Invent.WeaponEqpObjIndex = ObjIndex Then
492                                 Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                                Else
494                                 Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)

                                End If
                            
496                         Case HACHA_LEÑADOR
                            
                                ' Lo tiene equipado?
498                             If .Invent.WeaponEqpObjIndex = ObjIndex Then
500                                 Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Talar)
                                Else
502                                 Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                End If
                            
504                         Case PIQUETE_MINERO
                        
                                ' Lo tiene equipado?
506                             If .Invent.WeaponEqpObjIndex = ObjIndex Then
508                                 Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                                Else
510                                 Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                End If
                            
                        End Select

                    End If
            
                Case eOBJType.otLibroGuild
                    
                    If .GuildIndex = 0 Then
                        Call WriteConsoleMsg(UserIndex, "¡¿A que clan quieres dar Experiencia si no posees ninguno?!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Deshonrarás a tu clan si utilizas el Libro de Liderazgo", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    If GuildsInfo(.GuildIndex).Lvl = MAX_GUILD_LEVEL Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Tu clan ya ha alcanzado el máximo nivel!!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                        Call WriteConsoleMsg(UserIndex, "¡Debes estar en Zona Segura para utilizar el Libro!", FontTypeNames.FONTTYPE_INFORED)
                        Exit Sub

                    End If
    
                    Call mGuilds.Guilds_AddExp(UserIndex, Obj.GuildExp)
                    
                    ' Remove Object
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UpdateUserInv(False, UserIndex, Slot)

512             Case eOBJType.otTravel

514                 If .flags.Muerto = 1 Then
516                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! No puedes viajar en este estado.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
518                 If .flags.TargetNPC = 0 Then
520                     Call WriteConsoleMsg(UserIndex, "Debes hacer clic sobre '" & GetVar(Npcs_FilePath, "NPC" & ObjData(ObjIndex).RequiredNpc, "NAME") & "'", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                
522                 If Npclist(.flags.TargetNPC).numero <> ObjData(ObjIndex).RequiredNpc Then
524                     Call WriteConsoleMsg(UserIndex, "Debes hacer clic sobre '" & GetVar(Npcs_FilePath, "NPC" & ObjData(ObjIndex).RequiredNpc, "NAME") & "'", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                
526                 If Not MapInfo(Obj.TelepMap).CanTravel Then
528                     Call WriteConsoleMsg(UserIndex, "No puedes viajar al destino ¡Será mejor que corras!", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If

530                 .Counters.Shield = 3
532                 Call FindLegalPos(UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY)
534                 Call WarpUserChar(UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, True)
                
536                 Call WriteConsoleMsg(UserIndex, "Has llegado a tu destino.", FontTypeNames.FONTTYPE_INFOGREEN)
                
538                 Call QuitarUserInvItem(UserIndex, Slot, 1)
                
540                 Call UpdateUserInv(False, UserIndex, Slot)
                
542                 Call RefreshCharStatus(UserIndex)

544             Case eOBJType.otTransformVIP

546                 If .flags.Muerto = 1 Then
548                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar el skin estando vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
550                 If .flags.Mimetizado = 1 Or .flags.Transform Then
552                     Call WriteConsoleMsg(UserIndex, "Ya tienes un efecto de transformación.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                
554                 Call TransformVIP_User(UserIndex, Obj.Ropaje)
                
556             Case eOBJType.otBebidas

558                 If .flags.Muerto = 1 Then
560                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

562                 .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed

564                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
566                 .flags.Sed = 0
568                 Call WriteUpdateHungerAndThirst(UserIndex)
                
                    'Quitamos del inv el item
570                 Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                    ' Los admin invisibles solo producen sonidos a si mismos
572                 If .flags.AdminInvisible = 1 Then
574                     Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                    Else
576                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))

                    End If
                
578                 Call UpdateUserInv(False, UserIndex, Slot)
            
580             Case eOBJType.otLlaves

582                 If .flags.Muerto = 1 Then
584                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
586                 If .flags.TargetObj = 0 Then Exit Sub
588                 TargObj = ObjData(.flags.TargetObj)

                    '¿El objeto clickeado es una puerta?
590                 If TargObj.OBJType = eOBJType.otPuertas Then

                        '¿Esta cerrada?
592                     If TargObj.Cerrada = 1 Then

                            '¿Cerrada con llave?
594                         If TargObj.Llave > 0 Then
596                             If TargObj.clave = Obj.clave Then
                 
598                                 MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
600                                 .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
602                                 Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)

                                    Exit Sub

                                Else
604                                 Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)

                                    Exit Sub

                                End If

                            Else

606                             If TargObj.clave = Obj.clave Then
608                                 MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
610                                 Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
612                                 .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex

                                    Exit Sub

                                Else
614                                 Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)

                                    Exit Sub

                                End If

                            End If

                        Else
616                         Call WriteConsoleMsg(UserIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If

                    End If
            
618             Case eOBJType.otBotellaVacia

620                 If .flags.Muerto = 1 Then
622                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

624                 If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
626                     Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

628                 MiObj.Amount = 1
630                 MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
632                 Call QuitarUserInvItem(UserIndex, Slot, 1)

634                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
636                     Call TirarItemAlPiso(.Pos, MiObj)

                    End If
                
638                 Call UpdateUserInv(False, UserIndex, Slot)
            
640             Case eOBJType.otBotellaLlena

642                 If .flags.Muerto = 1 Then
644                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

646                 .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed

648                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
650                 .flags.Sed = 0
652                 Call WriteUpdateHungerAndThirst(UserIndex)
654                 MiObj.Amount = 1
656                 MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
658                 Call QuitarUserInvItem(UserIndex, Slot, 1)

660                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
662                     Call TirarItemAlPiso(.Pos, MiObj)

                    End If
                
664                 Call UpdateUserInv(False, UserIndex, Slot)
            
666             Case eOBJType.otPergaminos

668                 If .flags.Muerto = 1 Then
670                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
672                 If .Stats.MaxMan > 0 Then
674                     If .flags.Hambre = 0 And .flags.Sed = 0 Then
676                         Call AgregarHechizo(UserIndex, Slot)
678                         Call UpdateUserInv(False, UserIndex, Slot)
                        Else
680                         Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
682                     Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)

                    End If

684             Case eOBJType.otMinerales

686                 If .flags.Muerto = 1 Then
688                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

690                 Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
                    
                Case eOBJType.otTeleportInvoker
                    Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, TeleportInvoker)
                         
692             Case eOBJType.otInstrumentos

694                 If .flags.Muerto = 1 Then
696                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
698                 If Obj.Real Then '¿Es el Cuerno Real?
700                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
702                         If MapInfo(.Pos.Map).Pk = False Then
704                             ' Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)

                                '   Exit Sub

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
706                         If .flags.AdminInvisible = 1 Then
708                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))
                            Else
710                             Call AlertarFaccionarios(UserIndex)
712                             Call SendData(SendTarget.ToFaction, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

                            End If
                        
                            Exit Sub

                        Else
714                         Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If

716                 ElseIf Obj.Caos Then '¿Es el Cuerno Legión?

718                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
720                         If MapInfo(.Pos.Map).Pk = False Then
722                             Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)

                                Exit Sub

                            End If
                        
                            ' Los admin invisibles solo producen sonidos a si mismos
724                         If .flags.AdminInvisible = 1 Then
726                             Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))
                            Else
728                             Call AlertarFaccionarios(UserIndex)
730                             Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

                            End If
                        
                            Exit Sub

                        Else
732                         Call WriteConsoleMsg(UserIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If

                    End If

                    'Si llega aca es porque es o Laud o Tambor o Flauta
                    ' Los admin invisibles solo producen sonidos a si mismos
734                 If .flags.AdminInvisible = 1 Then
736                     Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))
                    Else
738                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

                    End If
               
740             Case eOBJType.otBarcos

742                 If .flags.Montando = 1 Then Exit Sub
                
744                 If ((LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) And .flags.Navegando = 0) Or .flags.Navegando = 1 Then
746                     Call DoNavega(UserIndex, Obj, Slot)
                    Else
748                     Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)

                    End If
                
750             Case eOBJType.otMonturas

752                 If .flags.Muerto = 1 Then
754                     Call WriteConsoleMsg(UserIndex, "¡¡No puedes montar tu mascota estando muerto!!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
756                 If ((LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False)) And .flags.Navegando = 0) Or .flags.Navegando = 1 Then
                        
758                     Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
                    Else
760                     Call DoEquita(UserIndex, Obj, Slot)

                    End If

            End Select
    
        End With

        '<EhFooter>
        Exit Sub

UseInvItem_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.UseInvItem " & "at line " & Erl

        '</EhFooter>
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        'If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        Call TirarTodosLosItems(UserIndex)
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarTodo. Error: " & Err.number & " - " & Err.description)
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ItemSeCae_Err
        '</EhHeader>

100     With ObjData(Index)
102         ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And (.Caos <> 1 Or .NoSeCae = 0) And .OBJType <> eOBJType.otLlaves And .OBJType <> eOBJType.otBarcos And .NoSeCae = 0
        End With

        '<EhFooter>
        Exit Function

ItemSeCae_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.ItemSeCae " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
    '***************************************************
    On Error GoTo ErrHandler

    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As Obj

    Dim ItemIndex As Integer

    Dim DropAgua  As Boolean
    
    With UserList(UserIndex)

        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex

                    DropAgua = True

                    ' Es Ladron?
                    If .Clase = eClass.Thief Then

                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 187 Then

                            ' Limitación por nivel, después dropea normalmente
                            If .Stats.Elv >= 40 Then
                                ' No dropea en agua
                                DropAgua = False
                            End If
                        End If
                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If

        Next i

    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en TirarTodosLosItems. Error: " & Err.number & " - " & Err.description)
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ItemNewbie_Err
        '</EhHeader>

100     If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
102     ItemNewbie = ObjData(ItemIndex).Newbie = 1
        '<EhFooter>
        Exit Function

ItemNewbie_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.ItemNewbie " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo TirarTodosLosItemsNoNewbies_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 23/11/2009
        '07/11/09: Pato - Fix bug #2819911
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        Dim i         As Byte

        Dim NuevaPos  As WorldPos

        Dim MiObj     As Obj

        Dim ItemIndex As Integer
    
100     With UserList(UserIndex)

102         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
104         For i = 1 To UserList(UserIndex).CurrentInventorySlots
106             ItemIndex = .Invent.Object(i).ObjIndex

108             If ItemIndex > 0 Then
110                 If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
112                     NuevaPos.X = 0
114                     NuevaPos.Y = 0
                    
                        'Creo MiObj
116                     MiObj.Amount = .Invent.Object(i).Amount
118                     MiObj.ObjIndex = ItemIndex
                        'Pablo (ToxicWaste) 24/01/2007
                        'Tira los Items no newbies en todos lados.
120                     Tilelibre .Pos, NuevaPos, MiObj, True, True

122                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
124                         Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        End If
                    End If
                End If

126         Next i

        End With

        '<EhFooter>
        Exit Sub

TirarTodosLosItemsNoNewbies_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.TirarTodosLosItemsNoNewbies " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub TirarTodosLosItemsEnMochila(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo TirarTodosLosItemsEnMochila_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/09 (Budi)
        '***************************************************
        Dim i         As Byte

        Dim NuevaPos  As WorldPos

        Dim MiObj     As Obj

        Dim ItemIndex As Integer
    
100     With UserList(UserIndex)

102         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
104         For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
106             ItemIndex = .Invent.Object(i).ObjIndex

108             If ItemIndex > 0 Then
110                 If ItemSeCae(ItemIndex) Then
112                     NuevaPos.X = 0
114                     NuevaPos.Y = 0
                    
                        'Creo MiObj
116                     MiObj.Amount = .Invent.Object(i).Amount
118                     MiObj.ObjIndex = ItemIndex
120                     Tilelibre .Pos, NuevaPos, MiObj, True, True

122                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
124                         Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        End If
                    End If
                End If

126         Next i

        End With

        '<EhFooter>
        Exit Sub

TirarTodosLosItemsEnMochila_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.TirarTodosLosItemsEnMochila " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo getObjType_Err
        '</EhHeader>

100     If ObjIndex > 0 Then
102         getObjType = ObjData(ObjIndex).OBJType
        End If
    
        '<EhFooter>
        Exit Function

getObjType_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.getObjType " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub moveItem(ByVal UserIndex As Integer, _
                    ByVal originalSlot As Integer, _
                    ByVal newSlot As Integer)
        '<EhHeader>
        On Error GoTo moveItem_Err
        '</EhHeader>

        Dim tmpObj      As UserOBJ

        Dim newObjIndex As Integer, originalObjIndex As Integer

100     If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

102     With UserList(UserIndex)

104         If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
106         If .flags.Comerciando Then Exit Sub
    
108         tmpObj = .Invent.Object(originalSlot)
110         .Invent.Object(originalSlot) = .Invent.Object(newSlot)
112         .Invent.Object(newSlot) = tmpObj
    
            'Viva VB6 y sus putas deficiencias.
114         If .Invent.AnilloEqpSlot = originalSlot Then
116             .Invent.AnilloEqpSlot = newSlot
118         ElseIf .Invent.AnilloEqpSlot = newSlot Then
120             .Invent.AnilloEqpSlot = originalSlot
            End If
    
122         If .Invent.AuraEqpSlot = originalSlot Then
124             .Invent.AuraEqpSlot = newSlot
126         ElseIf .Invent.AuraEqpSlot = newSlot Then
128             .Invent.AuraEqpSlot = originalSlot
            End If
    
130         If .Invent.ArmourEqpSlot = originalSlot Then
132             .Invent.ArmourEqpSlot = newSlot
134         ElseIf .Invent.ArmourEqpSlot = newSlot Then
136             .Invent.ArmourEqpSlot = originalSlot
            End If
    
138         If .Invent.BarcoSlot = originalSlot Then
140             .Invent.BarcoSlot = newSlot
142         ElseIf .Invent.BarcoSlot = newSlot Then
144             .Invent.BarcoSlot = originalSlot
            End If
    
146         If .Invent.CascoEqpSlot = originalSlot Then
148             .Invent.CascoEqpSlot = newSlot
150         ElseIf .Invent.CascoEqpSlot = newSlot Then
152             .Invent.CascoEqpSlot = originalSlot
            End If

154         If .Invent.EscudoEqpSlot = originalSlot Then
156             .Invent.EscudoEqpSlot = newSlot
158         ElseIf .Invent.EscudoEqpSlot = newSlot Then
160             .Invent.EscudoEqpSlot = originalSlot
            End If
    
162         If .Invent.MochilaEqpSlot = originalSlot Then
164             .Invent.MochilaEqpSlot = newSlot
166         ElseIf .Invent.MochilaEqpSlot = newSlot Then
168             .Invent.MochilaEqpSlot = originalSlot
            End If
    
170         If .Invent.MunicionEqpSlot = originalSlot Then
172             .Invent.MunicionEqpSlot = newSlot
174         ElseIf .Invent.MunicionEqpSlot = newSlot Then
176             .Invent.MunicionEqpSlot = originalSlot
            End If
    
178         If .Invent.WeaponEqpSlot = originalSlot Then
180             .Invent.WeaponEqpSlot = newSlot
182         ElseIf .Invent.WeaponEqpSlot = newSlot Then
184             .Invent.WeaponEqpSlot = originalSlot
            End If
    
186         If .Invent.MonturaSlot = originalSlot Then
188             .Invent.MonturaSlot = newSlot
190         ElseIf .Invent.MonturaSlot = newSlot Then
192             .Invent.MonturaSlot = originalSlot
            End If
    
194         If .Invent.ReliquiaSlot = originalSlot Then
196             .Invent.ReliquiaSlot = newSlot
198         ElseIf .Invent.ReliquiaSlot = newSlot Then
200             .Invent.ReliquiaSlot = originalSlot
            End If
    
202         If .Invent.MagicSlot = originalSlot Then
204             .Invent.MagicSlot = newSlot
206         ElseIf .Invent.MagicSlot = newSlot Then
208             .Invent.MagicSlot = originalSlot
            End If
            
                
            If .Invent.PendientePartySlot = originalSlot Then
                .Invent.PendientePartySlot = newSlot
            ElseIf .Invent.PendientePartySlot = newSlot Then
                .Invent.PendientePartySlot = originalSlot
            End If

210         Call UpdateUserInv(False, UserIndex, originalSlot)
212         Call UpdateUserInv(False, UserIndex, newSlot)
        End With

        '<EhFooter>
        Exit Sub

moveItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.moveItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub MoveItem_Bank(ByVal UserIndex As Integer, _
                         ByVal originalSlot As Integer, _
                         ByVal newSlot As Integer, _
                         ByVal TypeBank As Byte)
        '<EhHeader>
        On Error GoTo MoveItem_Bank_Err
        '</EhHeader>

        Dim tmpObj      As UserOBJ

        Dim newObjIndex As Integer, originalObjIndex As Integer

100     If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub
    
102     If TypeBank <> E_BANK.e_User And TypeBank <> E_BANK.e_Account Then Exit Sub
    
104     With UserList(UserIndex)

106         If (originalSlot > MAX_BANCOINVENTORY_SLOTS) Or (newSlot > MAX_BANCOINVENTORY_SLOTS) Then Exit Sub
        
108         Select Case TypeBank

                Case E_BANK.e_User
110                 tmpObj = .BancoInvent.Object(originalSlot)
112                 .BancoInvent.Object(originalSlot) = .BancoInvent.Object(newSlot)
114                 .BancoInvent.Object(newSlot) = tmpObj
                
116                 Call UpdateBanUserInv(False, UserIndex, originalSlot)
118                 Call UpdateBanUserInv(False, UserIndex, newSlot)
                
120             Case E_BANK.e_Account
122                 tmpObj = .Account.BancoInvent.Object(originalSlot)
124                 .Account.BancoInvent.Object(originalSlot) = .Account.BancoInvent.Object(newSlot)
126                 .Account.BancoInvent.Object(newSlot) = tmpObj
                
128                 Call UpdateBanUserInv_Account(False, UserIndex, originalSlot)
130                 Call UpdateBanUserInv_Account(False, UserIndex, newSlot)
            End Select
        
132         Call UpdateVentanaBanco(UserIndex)
    
        End With
        
        '<EhFooter>
        Exit Sub

MoveItem_Bank_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvUsuario.MoveItem_Bank " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
                   
