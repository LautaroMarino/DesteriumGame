Attribute VB_Name = "mRanking"
Option Explicit

' Cada Guardado de usuarios (Ejecutado cada 20 minutos y cada 180 en worldside), _
' en el bucle de usuarios llama al RankUser_Check(userindex) para ver si está en el ranking y actualizar las posiciones, o bien meterlo en el ranking
' Al finalizar el chequeo antes de cerrar el worldSive hago llamar al RankUser_Save() y actualizo el archivo también.

' Creado para PVP AO
' Creador: Lautaro Marino
' 23/10/2020


Public Const MAX_TOP As Byte = 50

Public Type tRank
    Name As String
    Points() As Long
    
    Elv As Byte
    Class As Byte
    Raze As Byte
    Promedy As Long
End Type


' Ranking General
Public Rank(1 To MAX_TOP) As tRank

Public Rank_Fight(1 To MAX_TOP) As tRank

Dim Manager As clsIniManager

Private FilePath_Rank As String

' Creación del archivo Ranking.DAT en caso de que no exista.
Public Sub Create_RankUsers()

On Error GoTo ErrHandler
    Dim intFile As Integer
    Dim A As Long, B As Long
    
    intFile = FreeFile
    
    Open FilePath_Rank For Output As #intFile
        
        ' Ranking Retos :: TOP 50
        Print #intFile, "[RETOS]"
        For A = 1 To MAX_TOP
            Print #intFile, A & "=-0-0-0"
        Next A
        
        ' Ranking General :: TOP 50
        Print #intFile, "[GENERAL]"
        For A = 1 To MAX_TOP
            Print #intFile, A & "=-0-0"
        Next A
        
        ' Ranking de Clases :: TOP 20
        For A = 1 To NUMCLASES
            Print #intFile, "[" & UCase$(ListaClases(A)) & "]"
            
            For B = 1 To MAX_TOP
                Print #intFile, B & "=-0-0"
            Next B
        Next A
        
        ' Información de los usuarios para el RENDER
        Print #intFile, "[RANKINFO]"
        For A = 0 To MAX_TOP_RENDER
            Print #intFile, "BODY" & A & "=0"
            Print #intFile, "HEAD" & A & "=0"
            Print #intFile, "HELM" & A & "=0"
            Print #intFile, "SHIELD" & A & "=0"
            Print #intFile, "WEAPON" & A & "=0"
            Print #intFile, "FACTION" & A & "=0"
        Next A
    
    Close #intFile
    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:Create_RankUsers: (" & Err.Number & ") " & Err.description)
End Sub

' Cargamos la información del Ranking
Public Sub Load_RankUsers()

On Error GoTo ErrHandler

    FilePath_Rank = DatPath & "RANKING.DAT"
    
    If Not FileExist(FilePath_Rank, vbArchive) Then Create_RankUsers
    
    Set Manager = New clsIniManager: Manager.Initialize FilePath_Rank
    
    Dim A As Long, B As Long
    Dim Temp As String
    
    ' Cargamos el TOP GENERAL
    For A = 1 To MAX_TOP
        ReDim Preserve Rank(A).Points(0) As Long
        Temp = Manager.GetValue("GENERAL", A)
        
        Rank(A).Name = ReadField(1, Temp, Asc("-"))
        Rank(A).Points(0) = val(ReadField(2, Temp, Asc("-")))
        Rank(A).Elv = val(ReadField(3, Temp, Asc("-")))
    Next A
    
    ' Cargamos el TOP RETOS
    For A = 1 To MAX_TOP
        ReDim Preserve Rank_Fight(A).Points(1) As Long
        Temp = Manager.GetValue("RETOS", A)
        
        Rank_Fight(A).Name = ReadField(1, Temp, Asc("-"))
        Rank_Fight(A).Points(0) = val(ReadField(2, Temp, Asc("-")))
        Rank_Fight(A).Points(1) = val(ReadField(3, Temp, Asc("-")))
        Rank_Fight(A).Elv = val(ReadField(4, Temp, Asc("-")))
    Next A
    
    ' Cargamos el TOP GENERAL: CLASES
    For A = 1 To NUMCLASES
        For B = 1 To MAX_TOP_CLASS
            ReDim Preserve Rank_Class(A, B).Points(0) As Long
            Temp = Manager.GetValue(ListaClases(A), B)
            
            ' Uso un array compuesto
            Rank_Class(A, B).Name = ReadField(1, Temp, Asc("-"))
            Rank_Class(A, B).Points(0) = val(ReadField(2, Temp, Asc("-")))
            Rank_Class(A, B).Elv = val(ReadField(3, Temp, Asc("-")))
        Next B
    Next A
    
    For A = 0 To MAX_TOP_RENDER
        With RankInfo(A)
            .Body = val(Manager.GetValue("RANKINFO", "BODY" & A))
            .Head = val(Manager.GetValue("RANKINFO", "HEAD" & A))
            .Helm = val(Manager.GetValue("RANKINFO", "HELM" & A))
            .Shield = val(Manager.GetValue("RANKINFO", "SHIELD" & A))
            .Weapon = val(Manager.GetValue("RANKINFO", "WEAPON" & A))
            .Faction = val(Manager.GetValue("RANKINFO", "FACTION" & A))
        
        End With
        
    Next A
        
    Set Manager = Nothing
    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:Load_RankUsers: (" & Err.Number & ") " & Err.description)
    
End Sub

' Guardamos la información del Ranking
Public Sub RankUser_Save()

On Error GoTo ErrHandler
    Set Manager = New clsIniManager: Manager.Initialize FilePath_Rank
    
    Dim A As Long, B As Long
    Dim Temp As String
    
    ' Guardamos el TOP GENERAL
    For A = 1 To MAX_TOP
        Call Manager.ChangeValue("GENERAL", A, Rank(A).Name & "-" & CStr(Rank(A).Points(0)) & "-" & CStr(Rank(A).Elv))
    Next A
    
    ' Guardamos el TOP RETOS
    For A = 1 To MAX_TOP
        Call Manager.ChangeValue("RETOS", A, Rank_Fight(A).Name & "-" & CStr(Rank_Fight(A).Points(0)) & "-" & CStr(Rank_Fight(A).Points(1)) & "-" & CStr(Rank(A).Elv))
    Next A
    
    ' Guardamos el TOP GENERAL: CLASES
    For A = 1 To NUMCLASES
        For B = 1 To MAX_TOP_CLASS
            Call Manager.ChangeValue(ListaClases(A), B, Rank_Class(A, B).Name & "-" & CStr(Rank_Class(A, B).Points(0)) & "-" & CStr(Rank_Class(A, B).Elv))
        Next B
    Next A
    
    ' Guardamos la información de los personajes para el RENDER
    For A = 0 To MAX_TOP_RENDER
        With RankInfo(A)
            Call Manager.ChangeValue("RANKINFO", "BODY" & A, CStr(.Body))
            Call Manager.ChangeValue("RANKINFO", "HEAD" & A, CStr(.Head))
            Call Manager.ChangeValue("RANKINFO", "HELM" & A, CStr(.Helm))
            Call Manager.ChangeValue("RANKINFO", "SHIELD" & A, CStr(.Shield))
            Call Manager.ChangeValue("RANKINFO", "WEAPON" & A, CStr(.Weapon))
            Call Manager.ChangeValue("RANKINFO", "FACTION" & A, CStr(.Faction))
        End With
    Next A
    
    Call Manager.DumpFile(FilePath_Rank)
        
    Set Manager = Nothing
    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:RankUser_Save: (" & Err.Number & ") " & Err.description)
    Set Manager = Nothing
End Sub


Public Sub RankUser_AddPoint(ByVal UserIndex As Integer, ByVal Value As Long)
    
    ' Agrega un Value negativo/positivo a los points del usuario
    
    With UserList(UserIndex)
        .Ranking.Points = .Ranking.Points + Value
        
        If .Ranking.Points <= 0 Then .Ranking.Points = 0
        
        Call WriteConsoleMsg(UserIndex, "Has recibido " & Value & " Puntos de Honor.", FontTypeNames.FONTTYPE_INFOGREEN)
         'Call WriteUpdatePoints(UserIndex)
    End With
    
End Sub

' Buscamos al personaje en el Ranking » ¿Está? Lo actualizamos. ¿No está? Lo intentamos agregar
' Este procedimiento se hace en cada guardado de usuarios (Cada 20 minutos aproximadamente, en TODOS los personajes ONLINE)
Public Sub RankUser_Check(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim Slot As Byte
    Dim DataRank As tRank
    
    With UserList(UserIndex)
        
        ReDim DataRank.Points(1) As Long
        
        ' Ranking General
        Slot = RankUser_Search(UCase$(.Name))
        
        If Slot > 0 Then
            Rank(Slot).Points(0) = .Ranking.Points
            Rank(Slot).Name = .Name
            Rank(Slot).Elv = .Stats.Elv
        Else
            DataRank.Name = .Name
            DataRank.Points(0) = .Ranking.Points
            DataRank.Elv = .Stats.Elv
            
            Call RankUser_CheckNew(DataRank)
        End If
        
        ' Ranking de Retos
        Slot = RankFight_Search(UCase$(.Name))
        
        If Slot > 0 Then
            Rank_Fight(Slot).Points(0) = .Stats.Retos1Ganados
            Rank_Fight(Slot).Points(1) = .Stats.Retos1Jugados
            Rank_Fight(Slot).Name = .Name
            Rank_Fight(Slot).Elv = .Stats.Elv
        Else
            DataRank.Name = .Name
            DataRank.Points(0) = .Stats.Retos1Ganados
            DataRank.Points(1) = .Stats.Retos1Jugados
            DataRank.Elv = .Stats.Elv
            
            Call RankFight_CheckNew(DataRank)
        End If
        
        ' Ranking General: Clases
        Slot = RankUser_Class_Search(UCase$(.Name), .Clase)
            
        If Slot > 0 Then
            Rank_Class(.Clase, Slot).Points(0) = .Ranking.Points
            Rank_Class(.Clase, Slot).Name = .Name
            Rank_Class(.Clase, Slot).Elv = .Stats.Elv
        Else
            DataRank.Name = .Name
            DataRank.Points(0) = .Ranking.Points
            DataRank.Elv = .Stats.Elv
            
            Call RankUser_Class_CheckNew(DataRank, .Clase)
                
        End If
        
        ' Actualizamos la información para el render
        Call Rank_CheckTop(UserIndex)

    End With
    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:RankUser_Check: (" & Err.Number & ") " & Err.description)
End Sub
Private Sub RankUser_CheckNew(ByRef DataRank As tRank)
    
On Error GoTo ErrHandler
    
    Dim A As Long, B As Long
    Dim Temp(1 To MAX_TOP) As tRank
    Dim Slot As Byte
    
    ' Buscamos una posición
    For A = 1 To MAX_TOP
        If DataRank.Points(0) > Rank(A).Points(0) Then
            Slot = A
            Exit For
        End If
    Next A
    
    If Slot > 0 Then
        ' Copia para no repetir
        For A = 1 To MAX_TOP
            Temp(A) = Rank(A)
        Next A
    
        ' Movemos +1 a los usuarios desde esta posición.
        For A = Slot To MAX_TOP - 1
            Rank(A + 1) = Temp(A)
        Next A
        
        Rank(Slot) = DataRank
    End If
    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:RankUser_CheckNew: (" & Err.Number & ") " & Err.description)
End Sub
Private Sub RankFight_CheckNew(ByRef DataRank As tRank)
    
On Error GoTo ErrHandler
    
    Dim A As Long, B As Long
    Dim Temp(1 To MAX_TOP) As tRank
    Dim TempRank As tRank
    Dim Slot As Byte
    
    ' Buscamos una posición
    For A = 1 To MAX_TOP
        If DataRank.Points(0) > Rank_Fight(A).Points(0) Then
            Slot = A
            Exit For
        End If
    Next A
    
    If Slot > 0 Then
        ' Copia para no repetir
        For A = 1 To MAX_TOP
            Temp(A) = Rank_Fight(A)
        Next A
    
        ' Movemos +1 a los usuarios desde esta posición.
        For A = Slot To MAX_TOP - 1
            Rank_Fight(A + 1) = Temp(A)
        Next A
        
        Rank_Fight(Slot) = DataRank
    End If
    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:RankUser_CheckNew: (" & Err.Number & ") " & Err.description)
End Sub

Private Sub RankUser_Class_CheckNew(ByRef DataRank As tRank, _
                                    ByVal UserClass As Byte)
                                    
On Error GoTo ErrHandler
    
    Dim A As Long, B As Long
    Dim Temp(1 To MAX_TOP_CLASS) As tRank
    Dim TempClass As tRank
    Dim Slot As Byte
    
    ' Buscamos una posición
    For A = 1 To MAX_TOP_CLASS
        If DataRank.Points(0) > Rank_Class(UserClass, A).Points(0) Then
            Slot = A
            Exit For
        End If
    Next A
    
    If Slot > 0 Then
        ' Copia para no repetir
        For A = 1 To MAX_TOP_CLASS
            Temp(A) = Rank_Class(UserClass, A)
        Next A
    
        ' Movemos +1 a los usuarios desde esta posición.
        For A = Slot To MAX_TOP_CLASS - 1
            Rank_Class(UserClass, A + 1) = Temp(A)
        Next A
        
        Rank_Class(UserClass, Slot) = DataRank
    End If

    
    Exit Sub
ErrHandler:
    Call LogError("ERROR Ranking:RankUser_Class_CheckNew: (" & Err.Number & ") " & Err.description)
End Sub





' FUNCIONES Y PROCEDIMIENTOS DE AYUDA ORGANIZACION



' Lista de NOMBRES :: Ranking de Clases
Public Function Rank_LisT_Users_Class(ByVal Class As eClass) As String
    
    Dim A As Long
    Dim Tmp As String
    
    For A = 1 To MAX_TOP_CLASS
        Tmp = Tmp & Rank_Class(Class, A).Name & SEPARATOR
    Next A
        
    If Len(Tmp) Then _
        Tmp = Left$(Tmp, Len(Tmp) - 1)
        
    Rank_LisT_Users_Class = Tmp
End Function

' Chequeamos si el personaje está en la posición n°1
Public Sub Rank_CheckTop(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        ' Ranking General
        If UCase$(.Name) = UCase$(Rank(1).Name) Then
            RankInfo(0).Body = .Char.Body
            RankInfo(0).Head = .Char.Head
            RankInfo(0).Helm = .Char.CascoAnim
            RankInfo(0).Shield = .Char.ShieldAnim
            RankInfo(0).Weapon = .Char.WeaponAnim
            RankInfo(0).Faction = .Faction.Status
            RankInfo(0).Name = .Name
        End If
        
        If UCase$(.Name) = UCase$(Rank_Class(.Clase, 1).Name) Then
            RankInfo(.Clase).Body = .Char.Body
            RankInfo(.Clase).Head = .Char.Head
            RankInfo(.Clase).Helm = .Char.CascoAnim
            RankInfo(.Clase).Shield = .Char.ShieldAnim
            RankInfo(.Clase).Weapon = .Char.WeaponAnim
            RankInfo(.Clase).Faction = .Faction.Status
            RankInfo(.Clase).Name = .Name
        End If
        
    
    End With
End Sub
