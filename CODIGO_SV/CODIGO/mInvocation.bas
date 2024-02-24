Attribute VB_Name = "mInvocation"
Option Explicit

' INVOCACIONES CON USUARIOS

Public Type tInvocaciones
    
    Activo As Byte
    
    'INFORMACION CARGADA
    Desc As String
    NpcIndex As Integer
    CantidadUsuarios As Byte
    mapa As Byte
    X() As Byte
    Y() As Byte
    
End Type

Public NumInvocaciones As Byte

Public Invocaciones()  As tInvocaciones

'[INIT]
'NumInvocaciones = 1

'[INVOCACION1] 'Mago del inframundo
'NpcIndex = 410

'Mapa = 1
'CantidadUsuarios = 2
'Pos1 = 40 - 60
'Pos2 = 70 - 80
Public Sub LoadInvocaciones()
        '<EhHeader>
        On Error GoTo LoadInvocaciones_Err
        '</EhHeader>
          
        Dim i        As Integer

        Dim X        As Integer

        Dim ln       As String
    
        Dim NpcIndex As Integer
    
        Dim Pos      As WorldPos
          
100     NumInvocaciones = val(GetVar(DatPath & "Invocaciones.dat", "INIT", "NumInvocaciones"))
          
102     ReDim Invocaciones(0 To NumInvocaciones) As tInvocaciones

104     For i = 1 To NumInvocaciones

106         With Invocaciones(i)
108             .Activo = 0
110             .CantidadUsuarios = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "CantidadUsuarios"))
112             .mapa = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Mapa"))
114             .NpcIndex = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "NpcIndex"))
116             .Desc = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Desc")
                      
118             ReDim .X(1 To .CantidadUsuarios)
120             ReDim .Y(1 To .CantidadUsuarios)
                      
122             For X = 1 To .CantidadUsuarios
124                 ln = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Pos" & X)
                          
126                 .X(X) = val(ReadField(1, ln, 45))
128                 .Y(X) = val(ReadField(2, ln, 45))
130             Next X
                
132             Pos.Map = .mapa
134             Pos.X = .X(1)
136             Pos.Y = .Y(1)
            
138             NpcIndex = SpawnNpc(.NpcIndex, Pos, False, False)
            
140             If NpcIndex Then
142                 .Activo = 1
144                 Npclist(NpcIndex).flags.Invocation = i
146                 Call UpdateInfoNpcs(Pos.Map, NpcIndex)
                End If
            
            End With

148     Next i
          
        '<EhFooter>
        Exit Sub

LoadInvocaciones_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mInvocation.LoadInvocaciones " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function InvocacionIndex(ByVal mapa As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte) As Byte
        '<EhHeader>
        On Error GoTo InvocacionIndex_Err
        '</EhHeader>

        Dim i As Integer

        Dim Z As Integer
          
100     InvocacionIndex = 0
          
        '// Devuelve el Index del mapa de invocación en el que está
102     For i = 1 To NumInvocaciones

104         With Invocaciones(i)

106             For Z = 1 To .CantidadUsuarios

108                 If .mapa = mapa And (.X(Z) = X) And .Y(Z) = Y Then
110                     InvocacionIndex = i

                        Exit For

                    End If

112             Next Z

            End With

114     Next i
              
        '<EhFooter>
        Exit Function

InvocacionIndex_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mInvocation.InvocacionIndex " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' if invocacacionindex = 0 then
Public Function PuedeSpawn(ByVal Index As Byte) As Boolean
        '<EhHeader>
        On Error GoTo PuedeSpawn_Err
        '</EhHeader>
          
        Dim Contador As Byte

        Dim i        As Integer
          
100     PuedeSpawn = False

102     For i = 1 To Invocaciones(Index).CantidadUsuarios

104         If MapData(Invocaciones(Index).mapa, Invocaciones(Index).X(i), Invocaciones(Index).Y(i)).UserIndex Then
106             Contador = Contador + 1
                  
108             If Contador = Invocaciones(Index).CantidadUsuarios Then
110                 PuedeSpawn = True
                End If
            End If

112     Next i
          
        '<EhFooter>
        Exit Function

PuedeSpawn_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mInvocation.PuedeSpawn " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PuedeRealizarInvocacion(ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo PuedeRealizarInvocacion_Err
        '</EhHeader>
100     PuedeRealizarInvocacion = False
          
102     With UserList(UserIndex)

104         If .flags.Muerto Then Exit Function
106         If .flags.Mimetizado Then Exit Function
108         If Not MapInfo(.Pos.Map).Pk Then Exit Function
        End With
          
110     PuedeRealizarInvocacion = True
        '<EhFooter>
        Exit Function

PuedeRealizarInvocacion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mInvocation.PuedeRealizarInvocacion " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub RealizarInvocacion(ByVal UserIndex As Integer, ByVal Index As Byte)
        '<EhHeader>
        On Error GoTo RealizarInvocacion_Err
        '</EhHeader>
          
        Dim Pos As WorldPos
          
        ' ¿Los usuarios están en las pos?
100     If PuedeSpawn(Index) Then
              
            Dim NpcIndex As Integer

102         Pos.Map = Invocaciones(Index).mapa
104         Pos.X = RandomNumber(Invocaciones(Index).X(1) - 3, Invocaciones(Index).X(1) + 3)
106         Pos.Y = RandomNumber(Invocaciones(Index).Y(1) - 3, Invocaciones(Index).Y(1) + 3)
              
108         FindLegalPos UserIndex, Pos.Map, Pos.X, Pos.Y
110         NpcIndex = SpawnNpc(Invocaciones(Index).NpcIndex, Pos, True, False)
              
112         If Not NpcIndex = 0 Then
114             Invocaciones(Index).Activo = 1
116             Npclist(NpcIndex).flags.Invocation = Index
118             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Invocaciones(Index).Desc, FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
          
        '<EhFooter>
        Exit Sub

RealizarInvocacion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mInvocation.RealizarInvocacion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

