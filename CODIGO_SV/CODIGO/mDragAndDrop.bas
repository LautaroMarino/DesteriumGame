Attribute VB_Name = "mDragAndDrop"

Option Explicit
 
Sub DragToUser(ByVal UserIndex As Integer, _
               ByVal tIndex As Integer, _
               ByVal Slot As Byte, _
               ByVal Amount As Integer)

    ' @ Author : maTih.-
    '            Drag un slot a un usuario.

    Dim tobj       As Obj

    Dim tString    As String

    Dim Espacio    As Boolean

    Dim ObjIndex   As Integer

    Dim errorfound As String

    'No quier el puto item

    On Error GoTo DragToUser_Error

    If Not CanDragObj(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex, errorfound) Then
        WriteConsoleMsg UserIndex, errorfound, FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If

    If UserList(UserIndex).flags.Comerciando Then Exit Sub

    If UserList(UserIndex).flags.Muerto = 1 Then
        WriteConsoleMsg UserIndex, "¡Estás Muerto!", FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If

    If UserList(tIndex).flags.Muerto = 1 Then
        WriteConsoleMsg UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If

    'Preparo el objeto.
    tobj.Amount = Amount
    tobj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
          
    If Amount < 1 Then

        Exit Sub

    End If

    If ObjData(tobj.ObjIndex).Plata = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes tirar los objetos Plata.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    If ObjData(tobj.ObjIndex).Oro = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes tirar los objetos Oro.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    
    
    If Not MeterItemEnInventario(tIndex, tobj) Then
        WriteConsoleMsg UserIndex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If

    
     ' Comprobamos si esta en una misión
    Call Quests_Check_Objs(UserIndex, tobj.ObjIndex, tobj.Amount)
                          
    'Quito el objeto.
    QuitarUserInvItem UserIndex, Slot, Amount
          
    'Hago un update de su inventario.
    UpdateUserInv False, UserIndex, Slot
          
    'Preparo el mensaje para userINdex (quien dragea)
          
    tString = "Le has arrojado"
          
    If tobj.Amount <> 1 Then
        tString = tString & " " & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
    Else
        tString = tString & " tu " & ObjData(tobj.ObjIndex).Name
    End If
          
    tString = tString & " a " & UserList(tIndex).Name
          
    'Envio el mensaje
    WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFO
          
    'Preparo el mensaje para el otro usuario (quien recibe)
    tString = UserList(UserIndex).Name & " te ha arrojado"
          
    If tobj.Amount <> 1 Then
        tString = tString & " " & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
    Else
        tString = tString & " su " & ObjData(tobj.ObjIndex).Name
    End If
          
    'Envio el mensaje al otro usuario
    WriteConsoleMsg tIndex, tString, FontTypeNames.FONTTYPE_INFO

    If ObjData(tobj.ObjIndex).Log = 1 Then
        Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eDropObj, "El personaje " & UserList(UserIndex).Name & " le ha arrojado a " & UserList(tIndex).Name & " el objeto: " & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name)
    End If
    
    On Error GoTo 0

    Exit Sub

DragToUser_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure DragToUser of Módulo MOD_DrAGDrOp in line " & Erl

End Sub
 
Public Sub DragToNPC(ByVal UserIndex As Integer, _
                     ByVal tNpc As Integer, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)
       
    ' @ Author : maTih.-
    '            Drag un slot a un npc.

    On Error GoTo DragToNPC_Error

    On Error GoTo ErrHandler
       
    Dim TeniaOro As Long

    Dim teniaObj As Integer

    Dim tmpIndex As Integer
       
    tmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    TeniaOro = UserList(UserIndex).Stats.Gld
    teniaObj = UserList(UserIndex).Invent.Object(Slot).Amount
       
    'Es un banquero?
    If UserList(UserIndex).flags.Comerciando Then Exit Sub

    'If tmpIndex < 1 Then
    'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
    'End If

    'If Amount < 1 Then
    'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
    'End If

    'If Amount < tmpIndex Then
    'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
    'End If
    If Amount > teniaObj Then
        WriteConsoleMsg UserIndex, "No tienes esa cantidad", FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If

    If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
        Call UserDejaObj(UserIndex, Slot, Amount)
        'No tiene más el mismo amount que antes? entonces depositó.

        If teniaObj <> UserList(UserIndex).Invent.Object(Slot).Amount Then
            WriteConsoleMsg UserIndex, "Has depositado " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_INFO
            UpdateUserInv False, UserIndex, Slot
        End If

        'Es un npc comerciante?
    ElseIf Npclist(tNpc).Comercia = 1 Then
        'El npc compra cualquier tipo de items?

        If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType Then
            Call Comercio(eModoComercio.Venta, UserIndex, tNpc, Slot, Amount, 0)
            'Ganó oro? si es así es porque lo vendió.

            If TeniaOro <> UserList(UserIndex).Stats.Gld Then
                WriteConsoleMsg UserIndex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_INFO
            End If

        Else
            WriteConsoleMsg UserIndex, "El npc no está interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_INFO
        End If
    End If
       
    Exit Sub
       
ErrHandler:

    On Error GoTo 0

    Exit Sub

DragToNPC_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure DragToNPC of Módulo MOD_DrAGDrOp in line " & Erl
 
End Sub
 
Public Sub DragToPos(ByVal UserIndex As Integer, _
                     ByVal X As Byte, _
                     ByVal Y As Byte, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)
       
    ' @ Author : maTih.-
    '            Drag un slot a una posición.
       
    Dim errorfound As String

    Dim tobj       As Obj

    Dim tString    As String
       
    'No puede dragear en esa pos?

    On Error GoTo DragToPos_Error

    If UserList(UserIndex).flags.Muerto = 1 Then
        WriteConsoleMsg UserIndex, "¡Estás Muerto!", FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If
    
    If UserList(UserIndex).flags.SlotEvent > 0 Then
        WriteConsoleMsg UserIndex, "¡No es posible utilizarlo en este tipo de eventos!", FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If
    
    If Not CanDragObj(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex, errorfound) Then
        WriteConsoleMsg UserIndex, errorfound, FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If

    If Not CanDragToPos(UserList(UserIndex).Pos.Map, X, Y, errorfound) Then
        WriteConsoleMsg UserIndex, errorfound, FontTypeNames.FONTTYPE_INFO

        Exit Sub

    End If
    
    'Creo el objeto.
    tobj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    tobj.Amount = Amount
    
    'Agrego el objeto a la posición.
    MakeObj tobj, UserList(UserIndex).Pos.Map, CInt(X), CInt(Y)
       
    'Quito el objeto.
    QuitarUserInvItem UserIndex, Slot, Amount
       
    'Actualizo el inventario
    UpdateUserInv False, UserIndex, Slot
       
    'Preparo el mensaje.
    tString = "¡Lanzas imprecisamente!"
       
    'If tobj.Amount <> 1 Then
    '      tString = tString & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
    'Else
    'tString = tString & "tu " & ObjData(tobj.ObjIndex).Name 'faltaba el tstring &
    ' End If
       
    'ENvio.
    WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFO

    If ObjData(tobj.ObjIndex).Log = 1 Then
        Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eDropObj, "El personaje " & UserList(UserIndex).Name & " draggeo el objeto: " & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name)
    End If
    
    On Error GoTo 0

    Exit Sub

DragToPos_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure DragToPos of Módulo MOD_DrAGDrOp in line " & Erl
       
End Sub
 
Private Function CanDragToPos(ByVal Map As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByRef error As String) As Boolean
       
    ' @ Author : maTih.-
    '            Devuelve si se puede dragear un item a x posición.
       
    On Error GoTo CanDragToPos_Error

    CanDragToPos = False

    'Zona segura?

    If Not MapInfo(Map).Pk Then
        error = "No está permitido arrojar objetos al suelo en zonas seguras."

        Exit Function

    End If
       
    'Ya hay objeto?

    If Not MapData(Map, X, Y).ObjInfo.ObjIndex = 0 Then
        error = "Hay un objeto en esa posición!"

        Exit Function

    End If
       
    'Tile bloqueado?

    If Not MapData(Map, X, Y).Blocked = 0 Then
        error = "No puedes arrojar objetos en esa posición"

        Exit Function

    End If
    
    If MapData(Map, X, Y).TileExit.Map <> 0 Then
        error = "¡Encontraste el limite del mapa!"

        Exit Function

    End If

    CanDragToPos = True

    On Error GoTo 0

    Exit Function

CanDragToPos_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure CanDragToPos of Módulo MOD_DrAGDrOp in line " & Erl
       
End Function
 
Private Function CanDragObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByRef error As String) As Boolean
       
    ' @ Author : maTih.-
    '            Devuelve si un objeto es drageable.
    On Error GoTo CanDragObj_Error

    CanDragObj = False
       
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData()) Then Exit Function

    If ObjData(ObjIndex).Newbie <> 0 Then
        error = "No puedes arrojar objetos newbies!"

        Exit Function

    End If
       
    'If ObjData(ObjIndex).VIP <> 0 Then
    'error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

    'Exit Function

    'End If
              
    If ObjData(ObjIndex).Real <> 0 Then
        error = "¡No puedes arrojar tus objetos faccionarios!"

        Exit Function

    End If
              
    If ObjData(ObjIndex).Caos <> 0 Then
        error = "¡No puedes arrojar tus objetos faccionarios!"

        Exit Function

    End If
    
    If ObjData(ObjIndex).OBJType = otBarcos And UserList(UserIndex).flags.Navegando > 0 Then
        error = "¡No puedes arrojar barcos si estas usando uno!"

        Exit Function

    End If
        
    CanDragObj = True

    On Error GoTo 0

    Exit Function

CanDragObj_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure CanDragObj of Módulo MOD_DrAGDrOp in line " & Erl
       
End Function
