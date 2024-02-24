Attribute VB_Name = "mFacciones"
Public Const PATH_FACTION As String = "\DAT\FACTION.DAT"

Public MAX_FACTION        As Byte

Public Enum eTipoFaction

    r_None = 0
    r_Armada = 1
    r_Caos = 2

End Enum

Public Type tRange

    Text As String
    Frags As Long
    Elv As Byte
    Gld As Long
    
    MinDef As Integer
    MaxDef As Integer

End Type

Public Type tFaction

    Status As eTipoFaction
    FragsCiu As Long
    FragsCri As Long
    FragsOther As Long
    Range As Byte
    
    StartDate As String
    StartElv As Byte
    StartFrags As Integer
    
    ExFaction As eTipoFaction

End Type

Public Type tInfoFaction

    Name As String
    TeamFaction As eTipoFaction
    AttackFaction As Byte
    TotalRange As Byte
    Range() As tRange

End Type

Public InfoFaction() As tInfoFaction

'Rangos de las facciones con sus requisitos
Public Sub LoadFactions()
        '<EhHeader>
        On Error GoTo LoadFactions_Err
        '</EhHeader>

        Dim Read As clsIniManager

        Dim A    As Long, B As Long

        Dim Temp As String
   
100     Set Read = New clsIniManager
   
102     Read.Initialize App.Path & PATH_FACTION
   
104     MAX_FACTION = val(Read.GetValue("INIT", "MAX_FACTION"))
   
106     ReDim InfoFaction(1 To MAX_FACTION) As tInfoFaction
   
108     For A = 1 To MAX_FACTION

110         With InfoFaction(A)
112             .Name = Read.GetValue("FACTION" & A, "Name")
          
114             .TeamFaction = val(Read.GetValue("FACTION" & A, "TeamFaction"))
116             .TotalRange = val(Read.GetValue("FACTION" & A, "TotalRange"))
118             .AttackFaction = val(Read.GetValue("FACTION" & A, "AttackFaction"))
          
120             ReDim .Range(0 To .TotalRange) As tRange
          
122             For B = 0 To .TotalRange
124                 Temp = Read.GetValue("FACTION" & A, "Range" & B)
              
126                 .Range(B).Text = ReadField(1, Temp, Asc("-"))
128                 .Range(B).Frags = val(ReadField(2, Temp, Asc("-")))
130                 .Range(B).Elv = val(ReadField(3, Temp, Asc("-")))
132                 .Range(B).Gld = val(ReadField(4, Temp, Asc("-")))
134                 .Range(B).MinDef = val(ReadField(5, Temp, Asc("-")))
136                 .Range(B).MaxDef = val(ReadField(6, Temp, Asc("-")))
138             Next B

            End With

140     Next A
   
142     Set Read = Nothing
   
        '<EhFooter>
        Exit Sub

LoadFactions_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.LoadFactions " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Agregamos el usuario a la facción correspondiente
Public Sub Faction_AddUser(ByVal UserIndex As Integer, ByVal Faction As eTipoFaction)
        '<EhHeader>
        On Error GoTo Faction_AddUser_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If Not Faction_CheckRequired(UserIndex, Faction, 0) Then Exit Sub
        
104         If .Faction.ExFaction > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Ya has pertenecido a una facción. Deberás pedir perdón para volver a ser miembro.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
        
108         If Faction = r_Armada Then
110             If .Faction.FragsCiu > 0 Then
112                 Call WriteConsoleMsg(UserIndex, "Has asesinado gente inocente. Deberás pedir perdón para volver a ser miembro.", FontTypeNames.FONTTYPE_WARNING)

                    Exit Sub

                End If
            End If
        
114         .Faction.Status = Faction
116         .Faction.Range = 0
        
118         Call Faction_RewardUser(UserIndex)
            'Call RankUser_AddPoint(UserIndex, 5)
120         WriteConsoleMsg UserIndex, "¡Te has enlistado! El rey ha decidido entregarte una armadura única para que te defienda al momento de defender su honor ¡Usala bien!", FontTypeNames.FONTTYPE_INFOGREEN
        
        End With

        '<EhFooter>
        Exit Sub

Faction_AddUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.Faction_AddUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Faction_RemoveUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Faction_RemoveUser_Err
        '</EhHeader>
    
100     With UserList(UserIndex).Faction
102         .ExFaction = .Status
104         .Status = 0
106         .Range = 0
108         .StartDate = vbNullString
110         .StartElv = 0
112         .StartFrags = 0
        
114         Call WriteConsoleMsg(UserIndex, "¡Facción removida!", FontTypeNames.FONTTYPE_INFO)
        
116         Call Guilds_CheckAlineation(UserIndex, a_Neutral)
        End With
    
        '<EhFooter>
        Exit Sub

Faction_RemoveUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.Faction_RemoveUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Function Faction_CheckRequired(ByVal UserIndex As Integer, _
                                       ByVal Faction As eTipoFaction, _
                                       ByVal FactionRange As Byte) As Boolean
        '<EhHeader>
        On Error GoTo Faction_CheckRequired_Err
        '</EhHeader>
    
        Dim Frags As Long
    
100     With UserList(UserIndex)

102         Select Case Faction

                Case eTipoFaction.r_Armada
104                 Frags = .Faction.FragsCri

106             Case eTipoFaction.r_Caos
108                 Frags = .Faction.FragsCiu
            End Select
        
110         If Frags < InfoFaction(Faction).Range(FactionRange).Frags Then
112             WriteConsoleMsg UserIndex, "Necesitas " & InfoFaction(Faction).Range(FactionRange).Frags & " Asesinados. Y tu tienes '" & Frags & "'.", FontTypeNames.FONTTYPE_INFO
114             Faction_CheckRequired = False

                Exit Function

            End If
        
116         If .Stats.Elv < InfoFaction(Faction).Range(FactionRange).Elv Then
118             WriteConsoleMsg UserIndex, "Mataste suficientes criminales, pero te faltan " & InfoFaction(Faction).Range(FactionRange).Elv - .Stats.Elv & " niveles para poder recibir la próxima recompensa.", FontTypeNames.FONTTYPE_INFO
120             Faction_CheckRequired = False

                Exit Function

            End If
        
122         If .Stats.Gld < InfoFaction(Faction).Range(FactionRange).Gld Then
124             WriteConsoleMsg UserIndex, "Necesitas " & InfoFaction(Faction).Range(FactionRange).Gld & " monedas de oro para poder recibir la próxima recompensa.", FontTypeNames.FONTTYPE_INFO
126             Faction_CheckRequired = False

                Exit Function

            End If

        End With
    
128     Faction_CheckRequired = True
        '<EhFooter>
        Exit Function

Faction_CheckRequired_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.Faction_CheckRequired " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Otorgamos armaduras iniciales
' Esto despues tiene que ir cargado desde .dat en el FACTION.DAT
Public Sub Faction_RewardUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Faction_RewardUser_Err
        '</EhHeader>
        Dim Obj As Obj
    
100     With UserList(UserIndex)
102         Obj.Amount = 1
        
104         If .Faction.Status = r_Armada Then
106             Select Case .Clase
                    Case eClass.Cleric, eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
108                     If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
110                         Obj.ObjIndex = 1046
                        Else
112                         Obj.ObjIndex = 1496
                        End If
                    
114                 Case eClass.Paladin, eClass.Warrior
116                     If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
118                         Obj.ObjIndex = 1044
                        Else
120                         Obj.ObjIndex = 1045
                        End If
                    
122                 Case eClass.Mage
124                     If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
126                         If .Genero = eGenero.Hombre Then
128                             Obj.ObjIndex = 1049
                            Else
130                             Obj.ObjIndex = 1048
                            End If
                        Else
132                         Obj.ObjIndex = 1047
                        End If
                End Select
            
            Else
                ' Legión Oscura

134             Select Case .Clase
                    Case eClass.Cleric, eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
136                     If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
138                         Obj.ObjIndex = 1057
                        Else
140                         Obj.ObjIndex = 1497
                        End If
                    
142                 Case eClass.Paladin, eClass.Warrior
144                     If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
146                         Obj.ObjIndex = 1055
                        Else
148                         Obj.ObjIndex = 1056
                        End If
                
150                 Case eClass.Mage
152                     If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
154                         If .Genero = eGenero.Hombre Then
156                             Obj.ObjIndex = 1060
                            Else
158                             Obj.ObjIndex = 1059
                            End If
                        Else
160                         Obj.ObjIndex = 1058
                        End If
                End Select
            
            End If
        
        
162         If Not MeterItemEnInventario(UserIndex, Obj) Then
164             Call Logs_User(.Name, eUser, eNone, "Enlistado en facción no le dió armadura inicial. Darsela manual")
            End If
        
        End With
        '<EhFooter>
        Exit Sub

Faction_RewardUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.Faction_RewardUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Asignamos un rango al personaje.
' Este procedimiento se llama cada vez que un usuario mata a alguien opuesto a su rango.
Public Sub Faction_CheckRangeUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Faction_CheckRangeUser_Err
        '</EhHeader>
    
        Dim Faction As eTipoFaction

        Dim Frags   As Long

100     With UserList(UserIndex)

102         If .Faction.Status = 0 Then
104             Call WriteConsoleMsg(UserIndex, "No perteneces a ninguna facción", FontTypeNames.FONTTYPE_INFO)

                Exit Sub
        
            End If
        
106         If (.Faction.Range) = InfoFaction(.Faction.Status).TotalRange Then Exit Sub
108         If Not Faction_CheckRequired(UserIndex, .Faction.Status, .Faction.Range + 1) Then Exit Sub
        
110         .Faction.Range = .Faction.Range + 1
        
112         If InfoFaction(.Faction.Status).Range(.Faction.Range).Gld > 0 Then
114             .Stats.Gld = .Stats.Gld - InfoFaction(.Faction.Status).Range(.Faction.Range).Gld
116             Call WriteUpdateGold(UserIndex)
            End If
        
            ' Ultimo Rango Flod Consola.
118         If .Faction.Range = InfoFaction(.Faction.Status).TotalRange Then
120             SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje " & .Name & " ha alcanzado el último rango de su facción. Felicitaciones", FontTypeNames.FONTTYPE_INFO)
            End If
            
        End With

        '<EhFooter>
        Exit Sub

Faction_CheckRangeUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.Faction_CheckRangeUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Comprobamos si un personaje puede atacar a otro.
' Las reglas son básicas. Usuarios de la misma facción se pueden atacar si la variable configurable AttackFaction=1
' Si son de distinta facción, podra atacar a la víctima si la variable TeamFaction no es igual al índice de facción del enemigo.
Public Function Faction_CanAttack(ByVal AttackerIndex As Integer, _
                                  ByVal VictimIndex As Integer)
        '<EhHeader>
        On Error GoTo Faction_CanAttack_Err
        '</EhHeader>

        Dim StatusAttacker As Byte

        Dim StatusVictim   As Byte
    
100     Faction_CanAttack = False
    
102     With UserList(AttackerIndex)
        
104         StatusAttacker = .Faction.Status
106         StatusVictim = UserList(VictimIndex).Faction.Status

108         If StatusAttacker = StatusVictim Then
110             If InfoFaction(StatusAttacker).AttackFaction > 0 Then
112                 Faction_CanAttack = True

                    Exit Function

                End If
        
            Else

114             If InfoFaction(StatusAttacker).TeamFaction <> StatusVictim Then
116                 Faction_CanAttack = True

                    Exit Function

                End If
        
            End If
        
        End With

        '<EhFooter>
        Exit Function

Faction_CanAttack_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mFacciones.Faction_CanAttack " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
