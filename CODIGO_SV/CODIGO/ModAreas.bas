Attribute VB_Name = "ModAreas"
Option Explicit

Public Const ENTITY_TYPE_PLAYER As Long = 0
Public Const ENTITY_TYPE_NPC    As Long = 1
Public Const ENTITY_TYPE_OBJECT As Long = 2

Public Const DEFAULT_ENTITY_WIDTH As Byte = 2
Public Const DEFAULT_ENTITY_HEIGHT As Byte = 2

Private World As Collision.Grid

Public Sub Initialise(ByVal Zones As Long)

    Set World = New Collision.Grid
    
    Call World.Initialise(Zones + 1)
    Call World.Attach(AddressOf OnCreateEntity, AddressOf OnDeleteEntity, AddressOf OnUpdateEntity)
    
End Sub

Public Sub CreateEntity(ByVal Name As Long, ByVal Tag As Long, ByRef Coordinates As WorldPos, ByVal Width As Byte, ByVal Height As Byte)
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    Dim RangoPlayer As Byte
    Dim RangoNpc As Byte
    
    UUID.Name = Name
    UUID.Type = Tag
    
    
    #If FullScreen = 1 Then
        RangoPlayer = 10
        RangoNpc = 10
    #Else
        RangoPlayer = 11
        RangoNpc = 11
    #End If
    
    
    Select Case Tag
        Case ENTITY_TYPE_PLAYER
            Call World.Create(UUID, RangoPlayer, Coordinates.Map, Coordinates.X, Coordinates.Y, Width, Height)    ' TODO: Width / Height / Radius
        Case ENTITY_TYPE_NPC
            Call World.Create(UUID, RangoNpc, Coordinates.Map, Coordinates.X, Coordinates.Y, Width, Height)     ' TODO: Width / Height / Radius
        Case ENTITY_TYPE_OBJECT
            Call World.Create(UUID, 0, Coordinates.Map, Coordinates.X, Coordinates.Y, Width, Height)   ' TODO: Width / Height
    End Select

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub CreateEntity de modAreas.bas")
End Sub

Public Sub DeleteEntity(ByVal Name As Long, ByVal Tag As Long)
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.Delete(UUID)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub DeleteEntity de modAreas.bas")
End Sub

Public Sub UpdateEntity(ByVal Name As Long, ByVal Tag As Long, ByRef Coordinates As WorldPos)
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.Update(UUID, Coordinates.X, Coordinates.Y)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub UpdateEntity de modAreas.bas")
End Sub

Public Function QueryEntities(ByVal Name As Long, ByVal Tag As Long, ByRef Result() As Collision.UUID, Optional ByVal Selection As Long = 255) As Long
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.Search(UUID, Selection, Result)
    
    QueryEntities = UBound(Result)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub QueryEntities de modAreas.bas")
End Function

Public Function QueryObservers(ByVal Name As Long, ByVal Tag As Long, ByRef Result() As Collision.UUID, Optional ByVal Selection As Long = 255) As Long
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.query(UUID, Selection, Result)
    
    QueryObservers = UBound(Result)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub QueryObservers de modAreas.bas")
End Function

Public Function Pack(ByVal Map As Long, ByVal X As Long, ByVal Y As Long) As Long

    Pack = ((Map And &H3FF) * &H4000) Or ((X And &H7F) * &H80) Or (Y And &H7F) ' 10 + 7 + 7 = 24B UniqueID
       
End Function

Public Function Unpack(ByVal ID As Long) As WorldPos
    Unpack.Map = (ID \ &H4000) And &H3FF
    Unpack.X = (ID \ &H80) And &H7F
    Unpack.Y = (ID And &H7F)
End Function

Private Sub OnCreateEntity(ByRef Instigator As Collision.UUID, ByRef Observer As Collision.UUID)
On Error GoTo ErrHandler

    Dim Coordinates As WorldPos
          
    If (Not Observer.Type = ENTITY_TYPE_PLAYER) Then
        Exit Sub
    End If

    'Debug.Print "OnCreateEntity (On Player)", Instigator.Name, Observer.Name

    Select Case Instigator.Type
        Case ENTITY_TYPE_PLAYER
            With UserList(Instigator.Name)
                If Not (.flags.AdminInvisible = 1) Then
                    If MakeUserChar(False, Observer.Name, Instigator.Name, .Pos.Map, .Pos.X, .Pos.Y) Then
                        If .flags.Navegando = 0 Then
                            If UserList(Observer.Name).flags.Privilegios And PlayerType.User Then
                                If (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura) Or (.flags.Invisible Or .flags.Oculto) Then
                                    Call WriteSetInvisible(Observer.Name, .Char.charindex, True)
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Case ENTITY_TYPE_NPC
            With Npclist(Instigator.Name)
                If (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.zonaOscura) Or ((UserList(Observer.Name).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call MakeNPCChar(False, Observer.Name, Instigator.Name, .Pos.Map, .Pos.X, .Pos.Y)
                End If
            End With
        Case ENTITY_TYPE_OBJECT
            Coordinates = Unpack(Instigator.Name)
            
            With ObjData(MapData(Coordinates.Map, Coordinates.X, Coordinates.Y).ObjInfo.ObjIndex)
                Call WriteObjectCreate(Observer.Name, MapData(Coordinates.Map, Coordinates.X, Coordinates.Y).ObjInfo.ObjIndex, .GrhIndex, Coordinates.X, Coordinates.Y, .Sound)
            End With
    End Select

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub OnCreateEntity de modAreas.bas")
End Sub

Private Sub OnDeleteEntity(ByRef Instigator As Collision.UUID, ByRef Observer As Collision.UUID)
On Error GoTo ErrHandler
    Dim Coordinates As WorldPos
            
    If (Not Observer.Type = ENTITY_TYPE_PLAYER) Then
        Exit Sub
    End If

    'Debug.Print "OnDeleteEntity (On Player)", Instigator.Name, Observer.Name

    Select Case Instigator.Type
        Case ENTITY_TYPE_PLAYER
            With UserList(Instigator.Name)
                If .flags.AdminInvisible <> 1 Then
                    Call SendData(SendTarget.ToOne, Observer.Name, PrepareMessageCharacterRemove(.Char.charindex))
                End If
            End With
        Case ENTITY_TYPE_NPC
            With Npclist(Instigator.Name)
                Call SendData(SendTarget.ToOne, Observer.Name, PrepareMessageCharacterRemove(.Char.charindex))
            End With
        Case ENTITY_TYPE_OBJECT
            Coordinates = Unpack(Instigator.Name)
            Call WriteObjectDelete(Observer.Name, Coordinates.X, Coordinates.Y)
    End Select

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub OnDeleteEntity de modAreas.bas")
End Sub

Private Sub OnUpdateEntity(ByRef Instigator As Collision.UUID, ByRef Observer As Collision.UUID)
On Error GoTo ErrHandler

    If (Not Observer.Type = ENTITY_TYPE_PLAYER) Then
        Exit Sub
    End If

    'Debug.Print "OnUpdateEntity (On Player)", Instigator.Name, Observer.Name

    Select Case Instigator.Type
        Case ENTITY_TYPE_PLAYER
            With UserList(Instigator.Name)
                If .flags.AdminInvisible <> 1 Then
                    Call SendData(SendTarget.ToOne, Observer.Name, PrepareMessageCharacterMove(.Char.charindex, .Pos.X, .Pos.Y))
                End If
            End With
        Case ENTITY_TYPE_NPC
            With Npclist(Instigator.Name)
                Call SendData(SendTarget.ToOne, Observer.Name, PrepareMessageCharacterMove(.Char.charindex, .Pos.X, .Pos.Y))
            End With
        Case ENTITY_TYPE_OBJECT
            ' IF AN OBJECT MOVE, THEN BURN THE SERVER BECAUSE IS HAUNTED
    End Select
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.description & ") en Sub OnUpdateEntity de modAreas.bas")
End Sub

