Attribute VB_Name = "mViajes"
' Creado por Matias y Santiago.
' 18/06/2019 :

Option Explicit

Public Const TRAVEL_MAX_NPCS As Byte = 50
Public Const TRAVEL_NPC      As Integer = 17

Public Const TRAVEL_NPC_HOME As Integer = 916


Private Function SearchRepeat(ByVal NpcIndex As Integer, ByRef Npcs() As tNpc) As Integer
        '<EhHeader>
        On Error GoTo SearchRepeat_Err
        '</EhHeader>

        Dim A As Long
    
100     SearchRepeat = -1

102     For A = LBound(Npcs) To UBound(Npcs)

104         If Npcs(A).NpcIndex = NpcIndex Then
106             SearchRepeat = A

                Exit Function

            End If

108     Next A
    
        '<EhFooter>
        Exit Function

SearchRepeat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mViajes.SearchRepeat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub MiniMap_SetChest(ByVal Map As Integer, ByVal ObjIndex As Integer)
    MiniMap(Map).ChestLast = MiniMap(Map).ChestLast + 1
    
    ReDim Preserve MiniMap(Map).Chest(1 To MiniMap(Map).ChestLast) As Integer
    MiniMap(Map).Chest(MiniMap(Map).ChestLast) = ObjIndex
End Sub
Public Sub MiniMap_SetInfo(ByVal Map As Integer)
    MiniMap(Map).Name = MapInfo(Map).Name
    MiniMap(Map).Pk = IIf(MapInfo(Map).Pk = True, 1, 0)
    MiniMap(Map).LvlMin = MapInfo(Map).LvlMin
    MiniMap(Map).LvlMax = MapInfo(Map).LvlMax
    MiniMap(Map).InviSinEfecto = MapInfo(Map).InviSinEfecto
    MiniMap(Map).CaenItem = MapInfo(Map).CaenItems
    MiniMap(Map).OcultarSinEfecto = MapInfo(Map).OcultarSinEfecto
    MiniMap(Map).InvocarSinEfecto = MapInfo(Map).InvocarSinEfecto
    MiniMap(Map).ResuSinEfecto = MapInfo(Map).ResuSinEfecto
    
    MiniMap(Map).Sub_Maps = MapInfo(Map).SubMaps
    
    If MiniMap(Map).Sub_Maps > 0 Then
        MiniMap(Map).Maps = MapInfo(Map).Maps
    End If
End Sub

Public Sub UpdateInfoNpcs(ByVal Map As Integer, ByVal NpcIndex As Integer)

    On Error GoTo ErrHandler

    Dim X          As Long, Y As Long, Z As Long

    Dim A          As Long, B As Long, c As Long

    Dim bkNpcs()   As Integer

    Dim Ultimo     As Integer

    Dim UltimoCopy As Integer

    Dim SlotRepeat As Integer
    
    SlotRepeat = SearchRepeat(Npclist(NpcIndex).Numero, MiniMap(Map).Npcs)
    
    If SlotRepeat = -1 Then
        MiniMap(Map).NpcsNum = MiniMap(Map).NpcsNum + 1
        
        With MiniMap(Map).Npcs(MiniMap(Map).NpcsNum)
            
            .Name = Npclist(NpcIndex).Name
     
            .Exp = Npclist(NpcIndex).GiveEXP
            .Gld = Npclist(NpcIndex).GiveGLD
            .NpcIndex = Npclist(NpcIndex).Numero
            'If .cant < 1000 Then
            '.cant = .cant + 1
            ' End If
        
            .Hp = Npclist(NpcIndex).Stats.MaxHp
            .MinHit = Npclist(NpcIndex).Stats.MinHit
            .MaxHit = Npclist(NpcIndex).Stats.MaxHit
            .Body = Npclist(NpcIndex).Char.Body
            .Head = Npclist(NpcIndex).Char.Head

            If .Hp > 0 Then
                'Debug.Print .Name & " NRO: " & Npclist(NpcIndex).Numero & " Vida: " & .Hp & " Experiencia: " & .Exp & " Oro: " & .Gld

            End If
            
            .NroItems = Npclist(NpcIndex).Invent.NroItems
            .NroDrops = Npclist(NpcIndex).NroDrops
            .Invent = Npclist(NpcIndex).Invent
            
            Dim Text As String
            
            'Text = vbNullString
            
            For B = 1 To .NroDrops
                .Drop(B) = Npclist(NpcIndex).Drop(B)
               ' Text = Text & vbCrLf & .Name & " NRO: " & Npclist(NpcIndex).Numero & " Drop: " & ObjData(Npclist(NpcIndex).Drop(B).ObjIndex).Name & " x" & Npclist(NpcIndex).Drop(B).Amount
                
            Next B
            
          '  Debug.Print Text
            
            For B = 1 To .NroItems
                .Invent.Object(B) = Npclist(NpcIndex).Invent.Object(B)
                'Text = Text & vbCrLf & .Name & " NRO: " & Npclist(NpcIndex).Numero & " Drop: " & ObjData(.Invent.Object(B).ObjIndex).Name & " x" & Npclist(NpcIndex).Invent.Object(B).Amount
                
                
                If Map = 1 Then
                    If Npclist(NpcIndex).Comercia = 1 Then
                        Call DataServer_AddObjSkin(.Invent.Object(B).ObjIndex)
                    End If
                End If
            Next B
            
        '    Text = Left$(Text, Len(Text) - 1)
           ' Debug.Print Text

            

            .NroSpells = Npclist(NpcIndex).flags.LanzaSpells
                
            If .NroSpells Then
                .Spells = Npclist(NpcIndex).Spells

            End If

        End With

    End If

    Exit Sub

    ' TESTEO
    For A = LBound(MiniMap(Map).Npcs) To UBound(MiniMap(Map).Npcs)

        With MiniMap(Map).Npcs(A)

            If .cant > 0 Then
                'Debug.Print .Name & " (x" & .cant & ")"
            
                For B = 1 To MAX_NPC_DROPS
                    Debug.Print "Drop" & B & ": " & .Drop(B).ObjIndex & "-" & .Drop(B).Amount
                Next B
            
            End If

        End With

    Next A
   
    Exit Sub

ErrHandler:

End Sub
