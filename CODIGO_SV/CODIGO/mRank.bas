Attribute VB_Name = "mRank"
Option Explicit

Public Const MAX_TOP       As Byte = 3
Public Const MAX_MONTH     As Byte = 12

Public Enum eRank

    eElv = 1
    eRetos1 = 2
    eTorneo = 3

End Enum

Public Type tRankChar

    Name As String
    Elv As Byte
    Class As Byte
    Value(1) As Long
    Promedy As Long

End Type

Public Type tRank

    Name As String
    
    Chars() As tRankChar

    Reset As Byte
    Max_Top_Users As Integer
    TempMonth(12) As tRankChar
    
End Type


Public Ranking_Month As Byte
Public Ranking(1 To MAX_TOP) As tRank

Private FilePath_Rank        As String

' Cargamos la información del Ranking
Public Sub Ranking_Load()

    On Error GoTo ErrHandler
    
    Dim Manager As clsIniManager
    
    FilePath_Rank = DatPath & "RANKING.DAT"
    
    If Not FileExist(FilePath_Rank, vbArchive) Then
        Ranking_Create
    End If
    
    Set Manager = New clsIniManager: Manager.Initialize FilePath_Rank
    
    Dim A    As Long, B As Long

    Dim Temp As String
    
    Ranking_Month = val(Manager.GetValue("INIT", "RANKING_LASTMONTH"))
    
    For A = 1 To MAX_TOP

        With Ranking(A)
            .Reset = val(Manager.GetValue(A, "RESET")) ' ¿Se reinicia o no?
            
            .Max_Top_Users = val(Manager.GetValue(A, "MAX_USERS"))
            
            ReDim .Chars(1 To .Max_Top_Users) As tRankChar
            
            For B = 1 To .Max_Top_Users
                Temp = Manager.GetValue(A, B)
                
                .Chars(B).Name = ReadField(1, Temp, Asc("-"))
                .Chars(B).Class = val(ReadField(2, Temp, Asc("-")))
                .Chars(B).Elv = val(ReadField(3, Temp, Asc("-")))
                .Chars(B).Value(0) = val(ReadField(4, Temp, Asc("-")))
                .Chars(B).Value(1) = val(ReadField(5, Temp, Asc("-")))
                
                
                If A = eRank.eElv Then
                    .Chars(B).Promedy = .Chars(B).Elv
                Else
                
                    If .Chars(B).Value(0) > .Chars(B).Value(1) Then
                         .Chars(B).Promedy = (.Chars(B).Value(0) - .Chars(B).Value(1))
                    Else
                         .Chars(B).Promedy = (.Chars(B).Value(1) - .Chars(B).Value(0))
                    End If
                   
                End If
                
                
            Next B
            
            For B = 1 To 12
                 Temp = Manager.GetValue(A, "MONTH" & B)
                
                .TempMonth(B).Name = ReadField(1, Temp, Asc("-"))
                .TempMonth(B).Class = val(ReadField(2, Temp, Asc("-")))
                .TempMonth(B).Elv = val(ReadField(3, Temp, Asc("-")))
                .TempMonth(B).Value(0) = val(ReadField(4, Temp, Asc("-")))
                .TempMonth(B).Value(1) = val(ReadField(5, Temp, Asc("-")))
                
            Next B
            
        End With

    Next A
        
    Set Manager = Nothing
    
    Exit Sub

ErrHandler:
    Set Manager = Nothing
    Call LogError("ERROR Ranking_Load: (" & Err.number & ") " & Err.description)
    
End Sub

' Chequea si hace falta reiniciar el Ranking
Public Sub Ranking_Loop_Reset()
        '<EhHeader>
        On Error GoTo Ranking_Loop_Reset_Err
        '</EhHeader>
    
        Dim Update As Boolean
    
        Dim A As Long
100     Dim Time As String: Time = Now
    
102     If Ranking_Month <> Month(Now) Then
104         Call Ranking_Save(True)
        End If
        '<EhFooter>
        Exit Sub

Ranking_Loop_Reset_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_Loop_Reset " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Ranking_ResetUsersAll(ByVal Rank As eRank)
        '<EhHeader>
        On Error GoTo Ranking_ResetUsersAll_Err
        '</EhHeader>

        Dim A As Long
        Dim NullChar As tRankChar
    
100     With Ranking(Rank)
    
102         For A = 1 To .Max_Top_Users
104             .Chars(A) = NullChar
106         Next A
        
        End With
        '<EhFooter>
        Exit Sub

Ranking_ResetUsersAll_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_ResetUsersAll " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
' Cargamos la información del Ranking
Public Sub Ranking_Save(Optional ByVal Reset As Boolean = False)
        '<EhHeader>
        On Error GoTo Ranking_Save_Err
        '</EhHeader>


        Dim Manager As clsIniManager
        Dim NullChar As tRankChar
        Dim CharWin As String
        Dim DateMonth As Integer
    
100     Set Manager = New clsIniManager
     
        If FileExist(FilePath_Rank, vbArchive) Then
            Manager.Initialize FilePath_Rank
        End If

        Dim A    As Long, B As Long

        Dim Temp As String

102     For A = 1 To MAX_TOP
104         With Ranking(A)
        
                ' Orden de los Top
106             Call mRank.Ranking_Ordenate(A)
            
                ' El Ranking se reinicia?
108             If (.Reset And Reset) Then
110                 CharWin = .Chars(1).Name & "-" & .Chars(1).Class & "-" & .Chars(1).Elv & "-" & .Chars(1).Value(0) & "-" & .Chars(1).Value(1)
                    
112                 Call Ranking_ResetUsersAll(A)
114                 Call Manager.ChangeValue(A, "MONTH" & Ranking_Month, CharWin)
                End If
            
            
                ' Guardado
116             For B = 1 To .Max_Top_Users
118                 Temp = .Chars(B).Name & "-" & .Chars(B).Class & "-" & .Chars(B).Elv & "-" & .Chars(B).Value(0) & "-" & .Chars(B).Value(1)
                
120                 Call Manager.ChangeValue(A, B, Temp)

122             Next B
            

            
            End With

       
124     Next A

126     If (Reset) Then
128         Ranking_Month = Month(Now)
130         Call Manager.ChangeValue("INIT", "RANKING_LASTMONTH", CStr(Ranking_Month))
        End If
        
132     Manager.DumpFile FilePath_Rank
    
134     Set Manager = Nothing
    
    
        '<EhFooter>
        Exit Sub

Ranking_Save_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_Save " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Ranking_Create()

    On Error GoTo ErrHandler

    Dim intFile As Integer

    Dim A       As Long, B As Long
    
    intFile = FreeFile
    
    Open FilePath_Rank For Output As #intFile
        
    For A = 1 To MAX_TOP
        Print #intFile, "[" & A & "]"
        
        Print #intFile, "RESETMONTH="
        Print #intFile, "MAX_USERS=50"
        
        For B = 1 To 50
            Print #intFile, B & "=-0-0-0-0"
        Next B
    Next A
    
    Close #intFile
    
    Exit Sub

ErrHandler:
    Call LogError("ERROR Ranking:Create_RankUsers: (" & Err.number & ") " & Err.description)
End Sub

' Comprueba si Existe el Usuario en los TOP
Public Sub Ranking_CheckExistUser(ByVal UserIndex As Integer)
    
    Dim A As Long
    
    
    With UserList(UserIndex)
        For A = 1 To MAX_TOP
            .Rank(A) = Ranking_ExistUser(UCase$(UserList(UserIndex).Name), A)
        Next A
        
    End With
    
End Sub

' Cambia el nick en el TOP necesario
Public Sub Ranking_Check_ChangeNick(ByVal UserIndex As Integer, ByVal UserName As String)
        '<EhHeader>
        On Error GoTo Ranking_Check_ChangeNick_Err
        '</EhHeader>
        Dim Rank As Integer
        Dim A As Long
        
100     For A = 1 To MAX_TOP
102         Rank = UserList(UserIndex).Rank(A)
        
104         If Rank > 0 Then
106             With Ranking(A)
108                 .Chars(Rank).Name = UCase$(UserName)
                End With
            End If
110     Next A
        
        '<EhFooter>
        Exit Sub

Ranking_Check_ChangeNick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_Check_ChangeNick " & _
               "at line " & Erl
       ' Resume Next
        '</EhFooter>
End Sub


' Comprueba si pasó la fecha estipulada en el user
Public Sub Ranking_CheckUser_Reset(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Ranking_CheckUser_Reset_Err
        '</EhHeader>
                                   
        Dim A As Long

100     With UserList(UserIndex)
    
            ' @ El usuario no tiene datos sobre la ultima fecha de RESET. Se la agregamos.
102         If .RankMonth <> Month(Now) Then
104             .RankMonth = Month(Now)
106             Call Ranking_ResetUser_Stats(UserIndex)
            Else
108             Call Ranking_CheckExistUser(UserIndex)
            End If
        
        End With
        '<EhFooter>
        Exit Sub

Ranking_CheckUser_Reset_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_CheckUser_Reset " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Se resetean las variables temporales del usuario
Private Sub Ranking_ResetUser_Stats(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        .Stats.Retos1Ganados = 0
        .Stats.Retos1Jugados = 0
        .Stats.DesafiosGanados = 0
        .Stats.DesafiosJugados = 0
    End With
    
End Sub

Public Function Ranking_CheckSlot(ByVal Rank As eRank, ByVal UserName As String) As Integer
        '<EhHeader>
        On Error GoTo Ranking_CheckSlot_Err
        '</EhHeader>
        Dim A As Long
    
100     For A = 1 To Ranking(Rank).Max_Top_Users
102         With Ranking(Rank).Chars(A)
104             If StrComp(UCase$(.Name), UserName) = 0 Then
106                 Ranking_CheckSlot = A
                    Exit Function
                End If
            End With
108     Next A
    
        '<EhFooter>
        Exit Function

Ranking_CheckSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_CheckSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Chequeamos un Slot TOP especifico.
Public Sub Ranking_CheckUser_Slot(ByVal UserIndex As Integer, _
                                  ByVal Selected As eRank)
    
    On Error GoTo ErrHandler
    Dim SlotUser As Integer: SlotUser = Ranking_CheckSlot(Selected, UCase$(UserList(UserIndex).Name))
    Dim A As Long
    
    If SlotUser > 0 Then
        Ranking(Selected).Chars(SlotUser) = Ranking_SetValues(UserIndex, Selected)
    Else
        
        Call Ranking_CheckNew(UserIndex, Selected, Ranking_SetValues(UserIndex, Selected))
    End If
    
    Exit Sub
ErrHandler:
End Sub


Private Sub Ranking_CheckNew(ByVal UserIndex As Integer, _
                             ByVal Rank As eRank, _
                             ByRef DataRank As tRankChar)
    
    On Error GoTo ErrHandler
    
    Dim A                        As Long, B As Long

    Dim Temp() As tRankChar

    Dim Slot                     As Integer
    
    Slot = Ranking_CheckNew_Values(UserIndex, Rank, DataRank)
    
    If Slot > 0 Then
        
        With Ranking(Rank)
            ReDim Temp(1 To .Max_Top_Users) As tRankChar
            
            ' Copia para no repetir
            For A = 1 To .Max_Top_Users
                Temp(A) = .Chars(A)
            Next A
        
            ' Movemos +1 a los usuarios desde esta posición.
            For A = Slot To .Max_Top_Users - 1
                .Chars(A + 1) = Temp(A)
            Next A
            
            .Chars(Slot) = DataRank
            UserList(UserIndex).Rank(Rank) = Slot
        End With
    End If
    
    Exit Sub

ErrHandler:
    Call LogError("ERROR Ranking_CheckNew: (" & Err.number & ") " & Err.description)
End Sub

' Remuve al personaje de todos los Rank en los que pertenezca. (BANEOS DE PERSONAJE, BORRADO DE PERSONAJES)
Public Function Ranking_RemoveUser_All(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo Ranking_RemoveUser_All_Err
        '</EhHeader>

        Dim A        As Long

        Dim SlotUser As Integer
    
100     For A = 1 To MAX_TOP
102         SlotUser = Ranking_ExistUser(UCase$(UserName), A)
        
104         If SlotUser > 0 Then
106             Call Ranking_RemoveUser(A, SlotUser)
            End If

108     Next A
    
        '<EhFooter>
        Exit Function

Ranking_RemoveUser_All_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_RemoveUser_All " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

'####
'#### GENERALES
'####

' Comprueba si un personaje está en el Ranking
Public Function Ranking_ExistUser(ByVal UserName As String, ByVal Rank As eRank) As Integer
        '<EhHeader>
        On Error GoTo Ranking_ExistUser_Err
        '</EhHeader>
    
        Dim A As Long
   
100     With Ranking(Rank)
102         For A = 1 To .Max_Top_Users

104             If StrComp(UCase$(.Chars(A).Name), UserName) = 0 Then
106                 Ranking_ExistUser = A

                    Exit Function

                End If
            
108         Next A
        End With

        '<EhFooter>
        Exit Function

Ranking_ExistUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_ExistUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Remueve un personaje del Ranking selecionado.
Public Sub Ranking_RemoveUser(ByVal Rank As eRank, ByVal SlotUser As Integer)
    
    With Ranking(Rank).Chars(SlotUser)
        .Name = vbNullString
        .Class = 0
        .Elv = 0
        .Value(0) = 0
        .Value(1) = 0
    End With

End Sub

' Setea los valores del personaje en una variable temporal
Public Function Ranking_SetValues(ByVal UserIndex As Integer, _
                                  ByVal Rank As eRank) As tRankChar

        '<EhHeader>
        On Error GoTo Ranking_SetValues_Err

        '</EhHeader>
    
        Dim Temp       As tRankChar

        Dim Promedy(1) As Long

        Dim Perdidos   As Long
        
100     Temp.Name = UserList(UserIndex).Name
102     Temp.Elv = UserList(UserIndex).Stats.Elv
104     Temp.Class = UserList(UserIndex).Clase
    
106     Select Case Rank

            Case eRank.eElv
                Temp.Value(0) = UserList(UserIndex).Stats.Exp
114             Temp.Value(1) = UserList(UserIndex).Stats.Elu
108             Temp.Promedy = Temp.Elv

110         Case eRank.eRetos1
                Perdidos = UserList(UserIndex).Stats.Retos1Jugados - UserList(UserIndex).Stats.Retos1Ganados
                Temp.Value(0) = UserList(UserIndex).Stats.Retos1Ganados
                Temp.Value(1) = UserList(UserIndex).Stats.Retos1Jugados
             
116             Temp.Promedy = (Temp.Value(1) - Perdidos)
        
118         Case eRank.eTorneo
120             Temp.Value(0) = UserList(UserIndex).Stats.Points
124             Temp.Promedy = Temp.Value(0)

        End Select

126     Ranking_SetValues = Temp
        '<EhFooter>
        Exit Function

Ranking_SetValues_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mRank.Ranking_SetValues " & "at line " & Erl

        

        '</EhFooter>
End Function

' Setea los valores del personaje en una variable temporal
Public Function Ranking_CheckNew_Values(ByVal UserIndex As Integer, _
                                        ByVal Rank As eRank, _
                                        ByRef DataRank As tRankChar) As Integer
        '<EhHeader>
        On Error GoTo Ranking_CheckNew_Values_Err
        '</EhHeader>
    
        Dim Temp As tRankChar

        Dim A    As Long
    
    
100     With Ranking(Rank)
102         For A = 1 To .Max_Top_Users
    
            
104             If .Chars(A).Promedy < DataRank.Promedy Then
106                 Ranking_CheckNew_Values = A
        
                    Exit Function
        
                End If
    
108         Next A
        
        End With
    
        '<EhFooter>
        Exit Function

Ranking_CheckNew_Values_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_CheckNew_Values " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Ordenamos la Lista del Ranking
Public Sub Ranking_Ordenate(ByVal Selected As eRank)
        '<EhHeader>
        On Error GoTo Ranking_Ordenate_Err
        '</EhHeader>

        Dim A    As Long, B As Long, c As Long

        Dim Temp As tRankChar
    
100         With Ranking(Selected)
102             For B = 1 To .Max_Top_Users
104                 For c = 1 To .Max_Top_Users - B
106                     If (.Chars(c).Promedy < .Chars(c + 1).Promedy) Then 'Or (Selected = eRank.eElv And .Chars(c).Elv = .Chars(c + 1).Elv And .Chars(c).Value(0) > .Chars(c + 1).Value(0))
108                         Temp = .Chars(c)
                            
110                         .Chars(c) = .Chars(c + 1)
112                         .Chars(c + 1) = Temp
                            
                        End If
114                 Next c
116             Next B
            End With

        '<EhFooter>
        Exit Sub

Ranking_Ordenate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_Ordenate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Lista de NOMBRES :: Ranking General
Public Function Ranking_List_Users(ByVal Rank As eRank) As String
        '<EhHeader>
        On Error GoTo Ranking_List_Users_Err
        '</EhHeader>
    
        Dim A   As Long

        Dim Tmp As String
    
100     For A = 1 To Ranking(Rank).Max_Top_Users
102         Tmp = Tmp & Ranking(Rank).Chars(A).Name & SEPARATOR
104     Next A
        
106     If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)
        
108     Ranking_List_Users = Tmp
        '<EhFooter>
        Exit Function

Ranking_List_Users_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRank.Ranking_List_Users " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub RankUser_AddPoint(ByVal UserIndex As Integer, ByVal Value As Long)

    With UserList(UserIndex)
        .Stats.Points = .Stats.Points + Value
    
    End With

End Sub
