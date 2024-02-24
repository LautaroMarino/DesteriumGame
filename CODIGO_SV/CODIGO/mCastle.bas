Attribute VB_Name = "mCastle"
Option Explicit

Public Enum eCastle
    CASTLE_NORTH = 1
    CASTLE_EAST = 2
    CASTLE_SOUTH = 3
    CASTLE_WEST = 4
    CASTLE_TOTAL = 5
End Enum

Public Type tCastle
    Name As String
    GuildName As String
    GuildIndex As Integer
    DateConquist As String
    
    LastSpam As Long
    
    Desc As String
    Map As Integer
    
    LastAttack As Integer
End Type

Public CastleBonus As Integer
Public CastleLast As Integer
Public Castle() As tCastle

' #Chequea si puede atacar al castillo
Public Function Castle_CanAttack(ByVal GuildIndex As Integer, ByVal CastleIndex As Integer) As Boolean
    
    On Error GoTo Errhandler
    
    With Castle(CastleIndex)
        If GuildIndex = .GuildIndex Then
            Castle_CanAttack = False
        End If
        
    End With
    
    Exit Function
Errhandler:
     
End Function

' # Comprueba que pueda aplicar el BONUS
Public Function Castle_CheckBonus(ByVal GuildIndex As Integer, ByVal CastleIndex As eCastle) As Boolean
    
    On Error GoTo Errhandler
    
    If GuildIndex > 0 Then
        If Castle(CastleIndex).GuildIndex = GuildIndex Then
            Castle_CheckBonus = True
        End If
    End If
                
    Exit Function
Errhandler:
End Function

' # Cargamos la info de castillos
Public Sub Castle_Load()

    On Error GoTo Errhandler
    
    Dim A As Long
    Dim FilePath As String
    
    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "Castle.dat"
    Manager.Initialize FilePath
    
    CastleLast = val(Manager.GetValue("INIT", "LAST"))
    CastleBonus = val(Manager.GetValue("INIT", "CASTLEBONUS"))
    
    ReDim Castle(1 To CastleLast) As tCastle
    
    For A = 1 To CastleLast
        With Castle(A)
            .Name = Manager.GetValue(A, "NAME")
            .GuildName = Manager.GetValue(A, "GUILDNAME")
            .DateConquist = Manager.GetValue(A, "DATE")
            .GuildIndex = val(Manager.GetValue(A, "GUILDINDEX"))
            .Desc = Manager.GetValue(A, "DESC")
            .LastAttack = val(Manager.GetValue(A, "LASTATTACK"))
            .Map = val(Manager.GetValue(A, "MAP"))
            
            
            MapInfo(.Map).FreeAttack = True
        End With
    Next A
    
    Set Manager = Nothing
    
    Castle_Save
    Exit Sub
Errhandler:
    
End Sub

' # Guardamos la info de castillos
Public Sub Castle_Save()

    On Error GoTo Errhandler
    
    Dim A As Long
    Dim FilePath As String
    
    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "Castle.dat"
    
    Call Manager.ChangeValue("INIT", "LAST", CastleLast)
    Call Manager.ChangeValue("INIT", "CASTLEBONUS", CastleBonus)
    
    For A = 1 To CastleLast
        With Castle(A)
            Call Manager.ChangeValue(A, "NAME", .Name)
            Call Manager.ChangeValue(A, "GUILDNAME", .GuildName)
            Call Manager.ChangeValue(A, "DATE", .DateConquist)
            Call Manager.ChangeValue(A, "GUILDINDEX", .GuildIndex)
            Call Manager.ChangeValue(A, "DESC", .Desc)
            Call Manager.ChangeValue(A, "LASTATTACK", .LastAttack)
            Call Manager.ChangeValue(A, "MAP", CStr(.Map))
        End With
    Next A
    
    Manager.DumpFile FilePath
    Set Manager = Nothing
    
    Exit Sub
Errhandler:
    
End Sub


' # Avisa que el castillo está siendo atacado.
Public Sub Castle_Attack(ByVal CastleIndex As Integer, ByVal GuildIndex As Integer)
    
    On Error GoTo Errhandler
    
    
    Dim Time As Long
    Time = GetTime
    
    With Castle(CastleIndex)
        
        .LastAttack = GuildIndex
        
        If (Time - .LastSpam) <= 60000 Then Exit Sub

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El '" & .Name & "' está siendo atacado por el clan '" & GuildsInfo(GuildIndex).Name & "'", FontTypeNames.FONTTYPE_GUILD))
        
        .LastSpam = GetTime
    End With
    
    Exit Sub
Errhandler:
    
End Sub

' # Chequea que tenga todos los Castillos Conquistados
Public Function Castle_CheckAllConquist(ByVal GuildIndex As Integer) As Boolean

    On Error GoTo Errhandler
    
    Dim A As Long
    
    For A = 1 To CastleLast
        With Castle(A)
            If .GuildIndex <> GuildIndex Then
                Exit Function
            End If
    
        End With
    Next A
    
    Castle_CheckAllConquist = True
    
    Exit Function
Errhandler:
    
End Function

Public Sub Castle_Close(ByVal CastleIndex As Integer)
    
    With Castle(CastleIndex)
        .DateConquist = 0
        .GuildIndex = 0
        .GuildName = 0
        .LastSpam = 0
        .LastAttack = 0
    
    End With
End Sub
' # Conquista el Castillo
Public Sub Castle_Conquist(ByVal CastleIndex As Integer, ByVal GuildIndex As Integer)

    On Error GoTo Errhandler
    
    
    With Castle(CastleIndex)
        .DateConquist = Format(Now, "dd/MM/yyyy hh:mm:ss")
        .GuildName = GuildsInfo(GuildIndex).Name
        .GuildIndex = GuildIndex
        .LastSpam = 0
        .LastAttack = 0
        CastleBonus = 0 ' # Reinicia la Fortaleza
        
        ' # Enviar a la BASE DE DATOS el ingreso
        Call WriteUpdateCastleConquist(.GuildName, CastleIndex, 1)
        
        ' # Envia un mensaje a discord
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El '" & .Name & "' pasa a estar en manos del clan '" & .GuildName & "' quien recibe: " & .Desc, FontTypeNames.FONTTYPE_GLOBAL))
    
        Dim TextDiscord As String
        
        TextDiscord = "El **'" & .Name & "'** pasa a estar en manos del clan **'" & .GuildName & "'** quien recibe: " & .Desc
        
        WriteMessageDiscord CHANNEL_CASTLE, TextDiscord
        
        ' # Check 4 Conquist Simultaneas
        If Castle_CheckAllConquist(GuildIndex) Then
            CastleBonus = GuildIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El clan '" & .GuildName & "' conquistó los 4 castillos. Recibe BONUS 10% de EXPERIENCIA y ORO.", FontTypeNames.FONTTYPE_GLOBAL))
            
            ' # Enviar a la base de DATOS.
            Call WriteUpdateCastleConquist(.GuildName, eCastle.CASTLE_TOTAL, 1)
            
            ' # Envia un mensaje a discord
            TextDiscord = "El clan **'" & .GuildName & "'** conquistó los 4 castillos. Recibe **BONUS** 10% de **EXPERIENCIA** y **ORO**."
            WriteMessageDiscord CHANNEL_CASTLE, TextDiscord
        End If
                
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(eSound.sVictory, 0, 0))
        Call Castle_Save
    End With
    
    
    Exit Sub
Errhandler:
    
End Sub


' # Viaja al Castillo en caso de que sea el DUEÑO
Public Sub Castle_Travel(ByVal UserIndex As Integer, ByVal CastleIndex As eCastle)

    On Error GoTo Errhandler
    
    Const Cost As Long = 15000
    
    With UserList(UserIndex)
    
        If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
        If .GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No puedes viajar al Castillo sin un CLAN y sin ser poseedor del Castillo!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
        If .GuildIndex <> Castle(CastleIndex).GuildIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes viajar al Castillo sin ser poseedor del Castillo.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
        
        If .Stats.Gld < Cost Then
            Call WriteConsoleMsg(UserIndex, "Necesitas " & Cost & " Monedas de Oro para poder viajar al Castillo.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
        
        ' #
        Time = GetTime
        
        If Not (Time - Castle(CastleIndex).LastSpam) <= 60000 Then
            Call WriteConsoleMsg(UserIndex, "El Castillo no está siendo atacado por nadie en un lapso de 60 segundos. NO puedes viajar directamente.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If

        .Stats.Gld = .Stats.Gld - Cost
        
        Call WriteUpdateGold(UserIndex)
        
        
        Dim Pos As WorldPos
        
        Pos.Map = Castle(CastleIndex).Map
        Pos.X = RandomNumber(20, 80)
        Pos.Y = RandomNumber(20, 80)
        
        Call EventWarpUser(UserIndex, Pos.Map, Pos.X, Pos.Y)
        Call WriteConsoleMsg(UserIndex, "¡Has viajado al castillo por un coste de " & Cost & " Monedas de Oro.", FontTypeNames.FONTTYPE_INFOGREEN)
        
    End With
    
    Exit Sub
Errhandler:

End Sub
