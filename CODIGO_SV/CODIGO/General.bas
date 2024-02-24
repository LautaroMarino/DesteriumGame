Attribute VB_Name = "General"
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

Public Running  As Boolean
Private Const MAX_TIME As Double = 2147483647 ' 2^31
Private LastTick As Double
Private overflowCount As Long
Global LeerNPCs As clsIniManager

Function DarCuerpoDesnudo_Genero(ByVal UserGenero As Byte, _
                                 ByVal UserRaza As Byte, _
                                 Optional ByVal Mimetizado As Boolean = False) As Integer

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/14/07
        'Da cuerpo desnudo a un usuario
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        '<EhHeader>
        On Error GoTo DarCuerpoDesnudo_Err

        '</EhHeader>

        Dim CuerpoDesnudo As Integer

102     Select Case UserGenero

            Case eGenero.Hombre

104             Select Case UserRaza

                    Case eRaza.Humano
106                     CuerpoDesnudo = 21

108                 Case eRaza.Drow
110                     CuerpoDesnudo = 32

112                 Case eRaza.Elfo
114                     CuerpoDesnudo = 21

116                 Case eRaza.Gnomo
118                     CuerpoDesnudo = 53

120                 Case eRaza.Enano
122                     CuerpoDesnudo = 53

                End Select

124         Case eGenero.Mujer

126             Select Case UserRaza

                    Case eRaza.Humano
128                     CuerpoDesnudo = 39

130                 Case eRaza.Drow
132                     CuerpoDesnudo = 40

134                 Case eRaza.Elfo
136                     CuerpoDesnudo = 39

138                 Case eRaza.Gnomo
140                     CuerpoDesnudo = 60

142                 Case eRaza.Enano
144                     CuerpoDesnudo = 60

                End Select

        End Select

        DarCuerpoDesnudo_Genero = CuerpoDesnudo
        '<EhFooter>
        Exit Function

DarCuerpoDesnudo_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.DarCuerpoDesnudo " & "at line " & Erl

        

        '</EhFooter>
End Function

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, _
                     Optional ByVal Mimetizado As Boolean = False)
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/14/07
        'Da cuerpo desnudo a un usuario
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        '<EhHeader>
        On Error GoTo DarCuerpoDesnudo_Err
        '</EhHeader>

        Dim CuerpoDesnudo As Integer

100     With UserList(UserIndex)

102         Select Case .Genero

                Case eGenero.Hombre

104                 Select Case .Raza

                        Case eRaza.Humano
106                         CuerpoDesnudo = 21

108                     Case eRaza.Drow
110                         CuerpoDesnudo = 32

112                     Case eRaza.Elfo
114                         CuerpoDesnudo = 21

116                     Case eRaza.Gnomo
118                         CuerpoDesnudo = 53

120                     Case eRaza.Enano
122                         CuerpoDesnudo = 53
                    End Select

124             Case eGenero.Mujer

126                 Select Case .Raza

                        Case eRaza.Humano
128                         CuerpoDesnudo = 39

130                     Case eRaza.Drow
132                         CuerpoDesnudo = 40

134                     Case eRaza.Elfo
136                         CuerpoDesnudo = 39

138                     Case eRaza.Gnomo
140                         CuerpoDesnudo = 60

142                     Case eRaza.Enano
144                         CuerpoDesnudo = 60
                    End Select
            End Select
          
146         If Mimetizado Then
148             .CharMimetizado.Body = CuerpoDesnudo
            Else
150             .Char.Body = CuerpoDesnudo
            End If
          
152         .flags.Desnudo = 1
        End With

        '<EhFooter>
        Exit Sub

DarCuerpoDesnudo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.DarCuerpoDesnudo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub Bloquear(ByVal toMap As Boolean, _
             ByVal sndIndex As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal B As Boolean)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        'b ahora es boolean,
        'b=true bloquea el tile en (x,y)
        'b=false desbloquea el tile en (x,y)
        'toMap = true -> Envia los datos a todo el mapa
        'toMap = false -> Envia los datos al user
        'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
        'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
        '***************************************************
        '<EhHeader>
        On Error GoTo Bloquear_Err
        '</EhHeader>

100     If toMap Then
102         Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, B))
        Else
104         Call WriteBlockPosition(sndIndex, X, Y, B)
        End If

        '<EhFooter>
        Exit Sub

Bloquear_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Bloquear " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo HayAgua_Err
        '</EhHeader>

100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then

102         With MapData(Map, X, Y)

104             If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520)) And .Graphic(2) = 0 Then
                
106                 HayAgua = True
                Else
108                 HayAgua = False
                End If

            End With

        Else
110         HayAgua = False
        End If

        '<EhFooter>
        Exit Function

HayAgua_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.HayAgua " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function HayLava(ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer) As Boolean
        '<EhHeader>
        On Error GoTo HayLava_Err
        '</EhHeader>

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        '***************************************************
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If (MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852) Or MapData(Map, X, Y).trigger = eTrigger.LavaActiva Then
104             HayLava = True
            Else
106             HayLava = False
            End If

        Else
108         HayLava = False
        End If

        '<EhFooter>
        Exit Function

HayLava_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.HayLava " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub LimpiarMundo()
        'SecretitOhs
        '<EhHeader>
        On Error GoTo LimpiarMundo_Err
        '</EhHeader>
100     'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))

        Dim MapaActual As Long

        Dim Y          As Long

        Dim X          As Long

        Dim bIsExit    As Boolean

102     For MapaActual = 1 To NumMaps

104         For Y = YMinMapSize To YMaxMapSize

106             For X = XMinMapSize To XMaxMapSize

108                 If MapData(MapaActual, X, Y).ObjInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 And MapInfo(MapaActual).Limpieza = 1 Then
110                     If (GetTime - MapData(MapaActual, X, Y).Protect) >= 60000 Then
                        
112                         If ItemNoEsDeMapa(MapaActual, X, Y, True) Then Call EraseObj(10000, MapaActual, X, Y)
                        End If
                    End If

114             Next X

116         Next Y

118     Next MapaActual

120     'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
        '<EhFooter>
        Exit Sub

LimpiarMundo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.LimpiarMundo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EnviarSpawnList_Err
        '</EhHeader>

        Dim K          As Long

        Dim npcNames() As String
    
100     ReDim npcNames(1 To UBound(SpawnList)) As String
    
102     For K = 1 To UBound(SpawnList)
104         npcNames(K) = SpawnList(K).NpcName
106     Next K
    
108     Call WriteSpawnList(UserIndex, npcNames())

        '<EhFooter>
        Exit Sub

EnviarSpawnList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EnviarSpawnList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Encripta fácilmente según la hora de la PC.
Private Function Encrypt_Value(ByVal Value As Long) As Long
    Encrypt_Value = Value Xor GetTime
End Function


Sub Main()

        '<EhHeader>
        On Error GoTo Main_Err

        '</EhHeader>

        Static variable As Integer

        '***************************************************
        'Author: Unknown
        'Last Modification: 15/03/2011
        '15/03/2011: ZaMa - Modularice todo, para que quede mas claro.
        '***************************************************

100     ChDir App.Path
102     ChDrive App.Path
    
104     GlobalActive = True
106     Call LoadMotd
108     Call BanIpCargar
110     Call AutoBan_Initialize
114     Call Challenge_SetMap
          
        PacketUseItem = ClientPacketID.UseItem
      
116     ReDim ListMails(0) As String
    
118     frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
        ' Start loading..
120     frmCargando.Show
    
        ' Constants & vars
122     frmCargando.Label1(2).Caption = "Cargando constantes..."
124     Call LoadConstants
126     DoEvents
    
        ' Arrays
128     frmCargando.Label1(2).Caption = "Iniciando Arrays..."
130     Call LoadArrays
    
        ' Cargamos base de datps
        Call MySql_Open
        
        ' Server.ini & Apuestas.dat
132     frmCargando.Label1(2).Caption = "Cargando Server.ini"
134     Call LoadSini
136     Call CargaApuestas
    
        ' Npcs_FilePath
138     frmCargando.Label1(2).Caption = "Cargando Criaturas"
140     Call CargaNpcsDat

        ' Objs_FilePath
142     frmCargando.Label1(2).Caption = "Cargando Objetos"
144     Call LoadOBJData

        ' Shop Items
        frmCargando.Label1(2).Caption = "Cargando Shop"
        Call Shop_Load
        Call Shop_Load_Chars
    
        ' Quests
146     Call LoadQuests
    
        ' Spell_FilePath
148     frmCargando.Label1(2).Caption = "Cargando Hechizos"
150     Call CargarHechizos
        
        ' Cargamos el mercado
152     Call mMao.Mercader_Load
    
        ' Balance.dat
162     frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
164     Call LoadBalance
    
        ' Animaciones
166     frmCargando.Label1(2).Caption = "Cargando Animaciones"
168     Call LoadAnimations
        
        ' Mapas
174     If BootDelBackUp Then
176         frmCargando.Label1(2).Caption = "Cargando BackUp"
178         Call CargarBackUp
        Else
180         frmCargando.Label1(2).Caption = "Cargando Mapas"
182         Call LoadMapData

        End If
        
        Call DataServer_Generate_ObjData
        
        Call Castle_Load
        
        ' Ruletas
        Call Ruleta_LoadItems
        
            ' Set de comerciantes en mapit
        Call Comerciantes_Load
        
        ' Generamos la info exportable al cliente de mapas
        Call DataServer_Generate_Maps
        
         ' Pathfinding
         Call InitPathFinding
         
        ' Eventos automáticos
184     Call LoadMapEvent
    
        ' Load Invocations
188     Call LoadInvocaciones
    
        ' Load Global Drops
190     Call Drops_Load
            
        ' Cargamos las facciones
196     Call LoadFactions
    
        ' Cargamos los RetoFast
198     Call LoadRetoFast
        
        ' Eventos AI
        Call Events_Load_PreConfig
        
        ' Pretorianos
        frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
        Call LoadPretorianData
    
        ' Map Sounds
200     Set SonidosMapas = New SoundMapInfo
202     Call SonidosMapas.LoadSoundMapInfo
    
        ' Home distance
204     Call generateMatrix(MATRIX_INITIAL_MAP)
    
        ' Connections
206     Call ResetUsersConnections
    
        ' Timers
208     Call InitMainTimers
    
        ' Sockets
210     Call SocketConfig
    
        'Call SocketConfig_Archive
    
        ' End loading..
212     Unload frmCargando
    
        'Log start time
214     LogServerStartTime
    
        'Ocultar
216     If HideMe = 1 Then
218         Call frmMain.InitMain(1)
        Else
220         Call frmMain.InitMain(0)

        End If
    
222     tInicioServer = GetTime
    
224     MercaderActivate = True
    
226     Running = True

228     While (Running)

230         Call Server.Poll
232         DoEvents
        Wend

        '<EhFooter>
        Exit Sub

Main_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.Main " & "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub DB_LoadSkills()
        '<EhHeader>
        On Error GoTo DB_LoadSkills_Err
        '</EhHeader>
        Dim Manager As clsIniManager
        Dim A As Long
    
100     Set Manager = New clsIniManager
        
102         Manager.Initialize DatPath & "skills.ini"
        
            ' Skills por Nivel que puede ganar el personaje
104         For A = 1 To 50
106             LevelSkill(A).LevelValue = val(Manager.GetValue("LEVELVALUE", "Lvl" & A))
108         Next A

            'NUMSKILLS = val(Manager.GetValue("INIT", "LastSkill"))
            'NUMSKILLSESPECIAL = val(Manager.GetValue("INIT", "LastSkillEspecial"))
            
          '  ReDim InfoSkill(1 To NUMSKILLS) As eInfoSkill
            
            ' Habilidades Cotidianas del Personaje
110         For A = 1 To NUMSKILLS
112             InfoSkill(A).Name = Manager.GetValue("SK" & A, "Name")
114             InfoSkill(A).MaxValue = val(Manager.GetValue("SK" & A, "MaxValue"))
116         Next A
    
           ' ReDim InfoSkillEspecial(1 To NUMSKILLSESPECIAL) As eInfoSkill
            
            ' Habilidades Extremas del Personaje
118         For A = 1 To NUMSKILLSESPECIAL
120             InfoSkillEspecial(A).Name = Manager.GetValue("SKESP" & A, "Name")
122             InfoSkillEspecial(A).MaxValue = val(Manager.GetValue("SKESP" & A, "MaxValue"))
124         Next A
        
126     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

DB_LoadSkills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.DB_LoadSkills " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub LoadConstants()
        '<EhHeader>
        On Error GoTo LoadConstants_Err
        '</EhHeader>

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Loads all constants and general parameters.
        '*****************************************************************
   
100     LastBackup = Format(Now, "Short Time")
102     Minutos = Format(Now, "Short Time")
    
        ' Paths
104     IniPath = App.Path & "\"
106     DatPath = App.Path & "\DAT\"
108     CharPath = App.Path & "\CHARS\CHARFILE\"
110     AccountPath = App.Path & "\CHARS\ACCOUNT\"
112     LogPath = App.Path & "\CHARS\LOGS\"
            
            
        ' Info Skills
          Call DB_LoadSkills
    
        ' Races
214     ListaRazas(eRaza.Humano) = "Humano"
216     ListaRazas(eRaza.Elfo) = "Elfo"
218     ListaRazas(eRaza.Drow) = "Drow"
220     ListaRazas(eRaza.Gnomo) = "Gnomo"
222     ListaRazas(eRaza.Enano) = "Enano"
    
        ' Classes
224     ListaClases(eClass.Mage) = "Mago"
226     ListaClases(eClass.Cleric) = "Clerigo"
228     ListaClases(eClass.Warrior) = "Guerrero"
230     ListaClases(eClass.Assasin) = "Asesino"
232     ListaClases(eClass.Thief) = "Ladron"
         
234     ListaClases(eClass.Bard) = "Bardo"
236     ListaClases(eClass.Druid) = "Druida"
238     ListaClases(eClass.Paladin) = "Paladin"
240     ListaClases(eClass.Hunter) = "Cazador"
242
          
        ' Attributes
274     ListaAtributos(eAtributos.Fuerza) = "Fuerza"
276     ListaAtributos(eAtributos.Agilidad) = "Agilidad"
278     ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
280     ListaAtributos(eAtributos.Carisma) = "Carisma"
282     ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
        ' Fishes
284     ListaPeces(1) = PECES_POSIBLES.PESCADO1
286     ListaPeces(2) = PECES_POSIBLES.PESCADO2
288     ListaPeces(3) = PECES_POSIBLES.PESCADO3
290     ListaPeces(4) = PECES_POSIBLES.PESCADO4
292     ListaPeces(5) = PECES_POSIBLES.PESCADO5
        
        #If Classic = 0 Then
            ListaPeces(6) = PECES_POSIBLES.PESCADO6
            ListaPeces(7) = PECES_POSIBLES.PESCADO7
        #End If
        
        'Bordes del mapa
294     MinXBorder = XMinMapSize + (XWindow \ 2)
296     MaxXBorder = XMaxMapSize - (XWindow \ 2)
298     MinYBorder = YMinMapSize + (YWindow \ 2)
300     MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
302     Set Denuncias = New cCola
304     Denuncias.MaxLenght = MAX_DENOUNCES


322         With Prision
324             .Map = 21
326             .X = 77
328             .Y = 15

            End With
    
330         With Libertad
332             .Map = 21
334             .X = 77
336             .Y = 29

            End With
            
            
338     MaxUsers = 0

340     Set aClon = New clsAntiMassClon
342     Set TrashCollector = New Collection

        '<EhFooter>
        Exit Sub

LoadConstants_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.LoadConstants " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub LoadArrays()
        '<EhHeader>
        On Error GoTo LoadArrays_Err
        '</EhHeader>

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Loads all arrays
        '*****************************************************************

        ' Load Records
100     Call LoadRecords
        ' Load guilds info
102     Call Guilds_Load
        ' Load spawn list
104     Call CargarSpawnList
        ' Load forbidden words
106     Call CargarForbidenWords
        ' Load Meditations
108     Call Meditation_LoadConfig
        ' Load Ranking
        'Call Load_RankUsers
        'Load Security
110     Call Initialize_Security
        ' Cargamos la pesca
112     Call Pesca_LoadItems
        ' Invasiones
114     Call Invations_Load
        ' Premiums Shop
        'Call Premiums_Load
        Call LoadHelp
        
116
        Call BotIntelligence_Load
118
        
        Call Arenas_Load
        Call MessageSpam_Load
        
        
        
        Call CargarFrasesOnFire
        '<EhFooter>
        Exit Sub

LoadArrays_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.LoadArrays " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub ResetUsersConnections()
        '<EhHeader>
        On Error GoTo ResetUsersConnections_Err
        '</EhHeader>

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Resets Users Connections.
        '*****************************************************************

        Dim LoopC As Long

100     For LoopC = 1 To MaxUsers
102         UserList(LoopC).ConnIDValida = False
104     Next LoopC
    
        '<EhFooter>
        Exit Sub

ResetUsersConnections_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.ResetUsersConnections " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub InitMainTimers()

        '<EhHeader>
        On Error GoTo InitMainTimers_Err

        '</EhHeader>

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Initializes Main Timers.
        '*****************************************************************

100     With frmMain
            .TimerGuardarUsuarios.Enabled = True
276         .TimerGuardarUsuarios.interval = IntervaloTimerGuardarUsuarios
102         .AutoSave.Enabled = True
104         .tPiqueteC.Enabled = True
106         .GameTimer.Enabled = True
108         .FX.Enabled = False
110         .Auditoria.Enabled = True
112         .KillLog.Enabled = True
114         .TIMER_AI.Enabled = True
            .tControlHechizos.Enabled = True
            .tControlHechizos.interval = 60000

        End With
    
        '<EhFooter>
        Exit Sub

InitMainTimers_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.InitMainTimers " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub SocketConfig()
        '<EhHeader>
        On Error GoTo SocketConfig_Err
        '</EhHeader>

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Sets socket config.
        '*****************************************************************

100     Set Writer = New Network.Writer
102     Set Server = New Network.Server
    
104     Call Server.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerSend, AddressOf OnServerReceive)
106     Call Server.Listen(MaxUsers, "0.0.0.0", Puerto)
    
108     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
        '<EhFooter>
        Exit Sub

SocketConfig_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.SocketConfig " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub LogServerStartTime()
        '<EhHeader>
        On Error GoTo LogServerStartTime_Err
        '</EhHeader>

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Logs Server Start Time.
        '*****************************************************************
        Dim N As Integer

100     N = FreeFile
102     Open LogPath & "Main.log" For Append Shared As #N
104     Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
106     Close #N

        '<EhFooter>
        Exit Sub

LogServerStartTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.LogServerStartTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
        '*****************************************************************
        'Se fija si existe el archivo
        '*****************************************************************
        '<EhHeader>
        On Error GoTo FileExist_Err
        '</EhHeader>

100     FileExist = LenB(dir$(File, FileType)) <> 0
        '<EhFooter>
        Exit Function

FileExist_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.FileExist " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
        '*****************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        'Gets a field from a delimited string
        '*****************************************************************
        '<EhHeader>
        On Error GoTo ReadField_Err
        '</EhHeader>

        Dim i          As Long

        Dim lastPos    As Long

        Dim CurrentPos As Long

        Dim delimiter  As String * 1
    
100     delimiter = Chr$(SepASCII)
    
102     For i = 1 To Pos
104         lastPos = CurrentPos
106         CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i
    
110     If CurrentPos = 0 Then
112         ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
        Else
114         ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
        End If

        '<EhFooter>
        Exit Function

ReadField_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.ReadField " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo MapaValido_Err
        '</EhHeader>

100     MapaValido = Map >= 1 And Map <= NumMaps
        '<EhFooter>
        Exit Function

MapaValido_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.MapaValido " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub MostrarNumUsers()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo MostrarNumUsers_Err
        '</EhHeader>

100     frmMain.txtNumUsers.Text = NumUsers
    
102     Call SendData(SendTarget.ToAll, 0, PrepareMessageUpdateOnline())
104     Call WriteUpdateOnline
        '<EhFooter>
        Exit Sub

MostrarNumUsers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.MostrarNumUsers " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



Sub Restart()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo Restart_Err
        '</EhHeader>


100     If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
        Dim LoopC As Long

102     For LoopC = 1 To MaxUsers
104         Protocol.Kick LoopC
        Next


106     ReDim UserList(1 To MaxUsers) As User
    
108     For LoopC = 1 To MaxUsers
110         UserList(LoopC).ConnIDValida = False
112     Next LoopC
    
114     LastUser = 0
116     NumUsers = 0
    
118     Call FreeNPCs
120     Call FreeCharIndexes
    
122     Call LoadSini
    
124     Call LoadOBJData
    
126     Call LoadMapData
    
128     Call CargarHechizos
    
130     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
        'Log it
        Dim N As Integer

132     N = FreeFile
134     Open LogPath & "Main.log" For Append Shared As #N
136     Print #N, Date & " " & Time & " servidor reiniciado."
138     Close #N
    
        'Ocultar
    
140     If HideMe = 1 Then
142         Call frmMain.InitMain(1)
        Else
144         Call frmMain.InitMain(0)
        End If

        '<EhFooter>
        Exit Sub

Restart_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Restart " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 15/11/2009
        '15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '**************************************************************
        '<EhHeader>
        On Error GoTo Intemperie_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
104             If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 2 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = True
            Else
106             Intemperie = False
            End If

        End With
    
        'En las arenas no te afecta la lluvia
108     If IsArena(UserIndex) Then Intemperie = False
        '<EhFooter>
        Exit Function

Intemperie_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Intemperie " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo TiempoInvocacion_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If .MascotaIndex > 0 Then
104             If Npclist(.MascotaIndex).Contadores.TiempoExistencia > 0 Then
106                 Npclist(.MascotaIndex).Contadores.TiempoExistencia = Npclist(.MascotaIndex).Contadores.TiempoExistencia - 1

108                 If Npclist(.MascotaIndex).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotaIndex, 0)
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

TiempoInvocacion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.TiempoInvocacion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EfectoTransformacion(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo EfectoTransformacion_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
102         .Stats.MinSta = .Stats.MinSta - 15

104         If .Stats.MinSta <= 0 Then .Stats.MinSta = 0
        
106         Call WriteUpdateSta(UserIndex)
        
108         If .Stats.MinSta <= 0 Then
110             Call Transform_User(UserIndex, 0)
112             Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoTransformacion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoTransformacion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Function CheckTunicaPolar(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)
        If .Invent.ArmourEqpObjIndex = 0 Then Exit Function
            
        
        If ObjData(.Invent.ArmourEqpObjIndex).AntiFrio > 0 Then
            CheckTunicaPolar = True 'ObjData(.Invent.ArmourEqpObjIndex).AntiFrio
        End If
    
    End With
    
End Function
Public Sub EfectoFrio(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo EfectoFrio_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unkonwn
        'Last Modification: 23/11/2009
        'If user is naked and it's in a cold map, take health points from him
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        Dim modifi As Integer
        Dim EfectoFrio As Integer
        
100     With UserList(UserIndex)
        
102         If .Counters.Frio < IntervaloFrio Then
104             .Counters.Frio = .Counters.Frio + 1
            Else

106             If MapInfo(.Pos.Map).Terreno = eTerrain.terrain_nieve Then
                    EfectoFrio = CheckTunicaPolar(UserIndex)
                     
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        EfectoFrio = CheckTunicaPolar(UserIndex)
                    End If
                    
108                 If Not CheckTunicaPolar(UserIndex) Then
110                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
112                     modifi = Porcentaje(.Stats.MaxHp, 10)
114                     .Stats.MinHp = .Stats.MinHp - modifi
                    
116                     If .Stats.MinHp < 1 Then
118                         Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
120                         .Stats.MinHp = 0
122                         Call UserDie(UserIndex)
                        Else
124                         Call WriteUpdateHP(UserIndex)
                        End If
                    End If
126             ElseIf MapInfo(.Pos.Map).Terreno = eTerrain.terrain_bosque Then
128                 If .flags.Desnudo = 1 Then
130                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
132                     modifi = Porcentaje(.Stats.MaxHp, 5)
134                     .Stats.MinHp = .Stats.MinHp - modifi
                    
                    
                    
136                     If .Stats.MinHp < 1 Then
138                         Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
140                         .Stats.MinHp = 0
142                         Call UserDie(UserIndex)
                        Else
144                         Call WriteUpdateHP(UserIndex)
                        End If
                
                    End If
                End If
            
146             .Counters.Frio = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoFrio_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoFrio " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo EfectoLava_Err
        '</EhHeader>

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 23/11/2009
        'If user is standing on lava, take health points from him
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
100     With UserList(UserIndex)

102         If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
104             .Counters.Lava = .Counters.Lava + 1
            Else

106             If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
108                 Call WriteConsoleMsg(UserIndex, "¡¡Quitate de la lava, te estás quemando!!", FontTypeNames.FONTTYPE_INFO)
110                 .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
                
112                 If .Stats.MinHp < 1 Then
114                     Call WriteConsoleMsg(UserIndex, "¡¡Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
116                     .Stats.MinHp = 0
118                     Call UserDie(UserIndex)
                    End If
                
120                 Call WriteUpdateHP(UserIndex)

                End If
            
122             .Counters.Lava = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoLava_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoLava " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EfectoAceleracion(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        .Counters.BuffoAceleration = .Counters.BuffoAceleration - 1
        
        If .Counters.BuffoAceleration <= 0 Then
            Call ActualizarVelocidadDeUsuario(UserIndex, False)
        End If
    
    End With
End Sub

''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el

'

Public Sub TravelingEffect(ByVal UserIndex As Integer)
        '******************************************************
        'Author: ZaMa
        'Last Update: 01/06/2010 (ZaMa)
        '******************************************************
        '<EhHeader>
        On Error GoTo TravelingEffect_Err
        '</EhHeader>
    
        Dim TiempoTranscurrido As Long
    

        ' Si ya paso el tiempo de penalizacion
100     If IntervaloGoHome(UserIndex) Then
102         Call HomeArrival(UserIndex)

        End If

        '<EhFooter>
        Exit Sub

TravelingEffect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.TravelingEffect " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo EfectoMimetismo_Err
        '</EhHeader>

        '******************************************************
        'Author: Unknown
        'Last Update: 16/09/2010 (ZaMa)
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
        '16/09/2010: ZaMa - Se recupera la apariencia de la barca correspondiente despues de terminado el mimetismo.
        '******************************************************
        Dim Barco As ObjData
    
100     With UserList(UserIndex)

102         If .flags.Transform > 0 Then Exit Sub
104         If .flags.SlotEvent > 0 Then Exit Sub
106         If .flags.TransformVIP > 0 Then Exit Sub
        
108         If .Counters.Mimetismo < IntervaloInvisible Then
110             .Counters.Mimetismo = .Counters.Mimetismo + 1
            Else
                'restore old char
112             Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
114             Call Mimetismo_Reset(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoMimetismo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoMimetismo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Mimetismo_Reset(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Mimetismo_Reset_Err
        '</EhHeader>

        Dim A As Long
        
100     With UserList(UserIndex)

102         If .flags.Navegando Then
104             If .flags.Muerto = 0 Then
106                 Call ToggleBoatBody(UserIndex)
                Else
108                 .Char.Body = iFragataFantasmal
110                 .Char.ShieldAnim = NingunEscudo
112                 .Char.WeaponAnim = NingunArma
114                 .Char.CascoAnim = NingunCasco

                      For A = 1 To MAX_AURAS
116                         .Char.AuraIndex(A) = NingunAura
                        Next A
                End If

            Else
118             .Char.Body = .CharMimetizado.Body
120             .Char.Head = .CharMimetizado.Head
122             .Char.CascoAnim = .CharMimetizado.CascoAnim
124             .Char.ShieldAnim = .CharMimetizado.ShieldAnim
126             .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                  
                  For A = 1 To MAX_AURAS
128                 .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
                  Next A
                  
            End If
            
130         With .Char
132             Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)
            End With
            
134         If .ShowName = False Then
136             .ShowName = True
138             Call RefreshCharStatus(UserIndex)
            End If
            
140         .Counters.Mimetismo = 0
142         .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
144         .flags.Ignorado = False
    
        End With

        '<EhFooter>
        Exit Sub

Mimetismo_Reset_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Mimetismo_Reset " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

'
Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 16/09/2010 (ZaMa)
        '16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
        '***************************************************
        '<EhHeader>
        On Error GoTo EfectoInvisibilidad_Err
        '</EhHeader>
        Dim TiempoTranscurrido As Long

100     With UserList(UserIndex)

102         If .Counters.Invisibilidad < IntervaloInvisible Then
104             .Counters.Invisibilidad = .Counters.Invisibilidad + 1
            
106             TiempoTranscurrido = (.Counters.Invisibilidad * frmMain.GameTimer.interval)
            
108             If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
110                 If TiempoTranscurrido = 40 Then
112                     Call WriteUpdateGlobalCounter(UserIndex, 1, ((IntervaloInvisible * 40) / 1000))
                    Else
114                     Call WriteUpdateGlobalCounter(UserIndex, 1, ((IntervaloInvisible * 40) / 1000) - ((.Counters.Invisibilidad * 40) / 1000))
                    End If
                End If
            
116             If .flags.Navegando = 0 Then
118                 Call EfectoInvisibilidad_Drawers(UserIndex)
                End If

            Else
120             .Counters.Invisibilidad = 0
122             .flags.Invisible = 0
            
124             Call WriteUpdateGlobalCounter(UserIndex, 1, 0)
            
126             If .flags.Oculto = 0 Then
128                 Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                    ' Si navega ya esta visible..
130                 If Not .flags.Navegando = 1 Then

                        'Si está en un oscuro no lo hacemos visible
132                     If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.zonaOscura Then
134                         Call SetInvisible(UserIndex, .Char.charindex, False)
                        End If
                    End If
                
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoInvisibilidad_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoInvisibilidad " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub EfectoInvisibilidad_Drawers(ByVal UserIndex As Integer)
        ' Author: Lautaro Marino
        ' Este procedimiento se encarga de mandar un paquete al cliente para visualizar los clientes invisibles durante un segundo.
        '<EhHeader>
        On Error GoTo EfectoInvisibilidad_Drawers_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

102         If .Counters.DrawersCount > 0 Then
104             .Counters.DrawersCount = .Counters.DrawersCount - 1
            
106             If .Counters.DrawersCount = 0 Then
108                 .Counters.Drawers = RandomNumberPower(7, 15)
110                 Call SetInvisible(UserIndex, .Char.charindex, .flags.Invisible, True)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, .flags.Invisible, True))
                End If
            
            End If
        
112         If .Counters.Drawers > 0 Then
114             .Counters.Drawers = .Counters.Drawers - 1
        
116             If .Counters.Drawers = 0 Then
118                 .Counters.DrawersCount = RandomNumberPower(1, 200)
120                 Call SetInvisible(UserIndex, .Char.charindex, .flags.Invisible, False)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, .flags.Invisible, False))
                End If
            End If
        
        End With

        '<EhFooter>
        Exit Sub

EfectoInvisibilidad_Drawers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoInvisibilidad_Drawers " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EfectoParalisisNpc_Err
        '</EhHeader>

100     With Npclist(NpcIndex)

102         If .Contadores.Paralisis > 0 Then
104             .Contadores.Paralisis = .Contadores.Paralisis - 1
            Else
106             .flags.Paralizado = 0
108             .flags.Inmovilizado = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoParalisisNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoParalisisNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EfectoCegueEstu_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If .Counters.Ceguera > 0 Then
104             .Counters.Ceguera = .Counters.Ceguera - 1
            Else

106             If .flags.Ceguera = 1 Then
108                 .flags.Ceguera = 0
110                 Call WriteBlindNoMore(UserIndex)
                End If

112             If .flags.Estupidez = 1 Then
114                 .flags.Estupidez = 0
116                 Call WriteDumbNoMore(UserIndex)
                End If
        
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoCegueEstu_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoCegueEstu " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/12/2010
        '02/12/2010: ZaMa - Now non-magic clases lose paralisis effect under certain circunstances.
        '***************************************************
        '<EhHeader>
        On Error GoTo EfectoParalisisUser_Err
        '</EhHeader>

        Dim TiempoTranscurrido As Long
    
100     With UserList(UserIndex)
    
102         If .Counters.Paralisis > 0 Then
        
                Dim CasterIndex As Integer

104             CasterIndex = .flags.ParalizedByIndex
        
                ' Only aplies to non-magic clases
106             If .Stats.MaxMan = 0 Then

                    ' Paralized by user?
108                 If CasterIndex <> 0 Then
                
                        ' Close? => Remove Paralisis
110                     If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
112                         Call RemoveParalisis(UserIndex)

                            Exit Sub
                        
                            ' Caster dead? => Remove Paralisis
114                     ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
116                         Call RemoveParalisis(UserIndex)

                            Exit Sub
                    
118                     ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then

                            ' Out of vision range? => Reduce paralisis counter
120                         If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                                ' Aprox. 1500 ms
122                             .Counters.Paralisis = IntervaloParalizadoReducido

                                Exit Sub

                            End If
                        End If
                
                        ' Npc?
                    Else
124                     CasterIndex = .flags.ParalizedByNpcIndex
                    
                        ' Paralized by npc?
126                     If CasterIndex <> 0 Then
                    
128                         If .Counters.Paralisis > IntervaloParalizadoReducido Then

                                ' Out of vision range? => Reduce paralisis counter
130                             If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                    ' Aprox. 1500 ms
132                                 .Counters.Paralisis = IntervaloParalizadoReducido

                                    Exit Sub

                                End If
                            End If
                        End If
                    
                    End If
                End If
            
134             .Counters.Paralisis = .Counters.Paralisis - 1
            
136             TiempoTranscurrido = (.Counters.Paralisis * frmMain.GameTimer.interval)
            
138             If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
140                     Call WriteUpdateGlobalCounter(UserIndex, 2, ((.Counters.Paralisis * 40) / 1000))
                End If

            Else
142             Call RemoveParalisis(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoParalisisUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoParalisisUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub RemoveParalisis(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo RemoveParalisis_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        'Removes paralisis effect from user.
        '***************************************************
100     With UserList(UserIndex)
102         .flags.Paralizado = 0
104         .flags.Inmovilizado = 0
106         .flags.ParalizedBy = vbNullString
108         .flags.ParalizedByIndex = 0
110         .flags.ParalizedByNpcIndex = 0
112         .Counters.Paralisis = 0
114         Call WriteParalizeOK(UserIndex)
        
116         WriteUpdateGlobalCounter UserIndex, 2, 0
        End With

        '<EhFooter>
        Exit Sub

RemoveParalisis_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.RemoveParalisis " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, _
                      ByRef EnviarStats As Boolean, _
                      ByVal Intervalo As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo RecStamina_Err

        '</EhHeader>

100     With UserList(UserIndex)

102         If .Pos.Map = 0 Then Exit Sub
106         If .flags.Transform > 0 Then Exit Sub

            Dim massta As Integer
        
108         If .flags.Desnudo Then
110             If .Stats.MinSta > 0 Then
112                 If .Counters.STACounter < Intervalo Then
114                     .Counters.STACounter = .Counters.STACounter + 1
                    Else
116                     EnviarStats = True
118                     .Counters.STACounter = 0

120                     massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
122                     .Stats.MinSta = .Stats.MinSta - massta
                    
124                     If .Stats.MinSta <= 0 Then
126                         .Stats.MinSta = 0

                        End If

                    End If

                End If

            Else
        
128             If .Stats.MinSta < .Stats.MaxSta Then
130                 If .Counters.STACounter < Intervalo Then
132                     .Counters.STACounter = .Counters.STACounter + 1
                    Else
134                     EnviarStats = True
136                     .Counters.STACounter = 0
                        'If .flags.Desnudo Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
                   
138                     massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 10))
140                     .Stats.MinSta = .Stats.MinSta + massta

142                     If .Stats.MinSta > .Stats.MaxSta Then
144                         .Stats.MinSta = .Stats.MaxSta

                        End If

                    End If

                End If

            End If

        End With
    
        '<EhFooter>
        Exit Sub

RecStamina_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.RecStamina " & "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub User_EfectoIncineracion(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo User_EfectoIncineracion_Err
        '</EhHeader>
        Dim N As Integer
    
100     With UserList(UserIndex)

102         If .Counters.Incinerado < IntervaloFrio Then
104             .Counters.Incinerado = .Counters.Incinerado + 1
            Else
106             .Counters.Incinerado = 0
            
108             Call WriteConsoleMsg(UserIndex, "¡Te estas incinerando, si no te curas morirás!", FontTypeNames.FONTTYPE_VENENO)
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sFogata, .Pos.X, .Pos.Y, .Char.charindex, True))
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FX_INCINERADO, -1))
            
114             N = RandomNumber(1, 50)
116             .Stats.MinHp = .Stats.MinHp - N

118             If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
120             Call WriteUpdateHP(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

User_EfectoIncineracion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.User_EfectoIncineracion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub Npc_EfectoIncineracion(ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo Npc_EfectoIncineracion_Err
        '</EhHeader>
        Dim N As Integer
        Dim UserIndex As Integer
    
100     With Npclist(NpcIndex)

102         If .Contadores.Incinerado < IntervaloFrio Then
104             .Contadores.Incinerado = .Contadores.Incinerado + 1
            Else
106             .Contadores.Incinerado = 0
            
108             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(eSound.sFogata, .Pos.X, .Pos.Y, .Char.charindex, True))
110             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FX_INCINERADO, -1))
            
112             N = RandomNumber(1, 50)
114             .Stats.MinHp = .Stats.MinHp - N
        
116             UserIndex = .Owner
118             If .Stats.MinHp < 1 Then Call MuereNpc(NpcIndex, UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

Npc_EfectoIncineracion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Npc_EfectoIncineracion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub EfectoVeneno(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EfectoVeneno_Err
        '</EhHeader>

        Dim N As Integer
    
100     With UserList(UserIndex)

102         If .Counters.Veneno < IntervaloVeneno Then
104             .Counters.Veneno = .Counters.Veneno + 1
            Else
106             Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
108             .Counters.Veneno = 0
110             N = RandomNumber(1, 5)
112             .Stats.MinHp = .Stats.MinHp - N

114             If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
116             Call WriteUpdateHP(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

EfectoVeneno_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EfectoVeneno " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

        '***************************************************
        'Author: ??????
        'Last Modification: 08/06/11 (CHOTS)
        'Le agregué que avise antes cuando se te está por ir
        '
        'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
        '***************************************************
        '<EhHeader>
        On Error GoTo DuracionPociones_Err

        '</EhHeader>

        Const SEGUNDOS_AVISO   As Byte = 5
        
        Dim Tick               As Long

        Dim TiempoTranscurrido As Long
        
        Tick = GetTime
        'CHOTS | Los segundos antes que se te acabe que te avisa

100     With UserList(UserIndex)

            'Controla la duracion de las pociones
102         If .flags.DuracionEfecto > 0 Then
104             .flags.DuracionEfecto = .flags.DuracionEfecto - 1
                    
108             TiempoTranscurrido = (.flags.DuracionEfecto * frmMain.GameTimer.interval)
            
                If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
                    Call WriteUpdateGlobalCounter(UserIndex, 3, .flags.DuracionEfecto / 40)

                End If
            
                If ((.flags.DuracionEfecto / 25) <= SEGUNDOS_AVISO) Then    'CHOTS | Lo divide por 25 por el intervalo del Timer (40x25=1000=1seg)
                    If Tick - .Counters.RuidoDopa > 5000 Then
                        .Counters.RuidoDopa = Tick
                        Call WriteStrDextRunningOut(UserIndex)

                    End If

110                 '  .flags.UltimoMensaje = 221
                End If

112             If .flags.DuracionEfecto = 0 Then
114                 ' .flags.UltimoMensaje = 222
116                 .flags.TomoPocion = False
118                 .flags.TipoPocion = 0

                    'volvemos los atributos al estado normal
                    Dim LoopX As Integer
                
120                 For LoopX = 1 To NUMATRIBUTOS
122                     .Stats.UserAtributos(LoopX) = .Stats.UserAtributosBackUP(LoopX)
124                 Next LoopX
                
126                 Call WriteUpdateStrenghtAndDexterity(UserIndex)

                End If

            End If
        
            Dim UpdateMAN As Boolean

            Dim UpdateHP  As Boolean, TempTick As Long
        
            ' Pociones Azules (Clic)
128         If .PotionBlue_Clic > 0 Then
130             .PotionBlue_Clic_Interval = .PotionBlue_Clic_Interval + 1
            
132             If .PotionBlue_Clic_Interval >= TOLERANCE_POTIONBLUE_CLIC Then
134                 .PotionBlue_Clic = .PotionBlue_Clic - 1
                
136                 If .PotionBlue_Clic > 0 Then
138                     .Stats.MinMan = .Stats.MinMan + Porcentaje(.Stats.MaxMan, 3) + .Stats.Elv \ 2 + 40 / .Stats.Elv
                                
140                     If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan
142                     UpdateMAN = True

                    End If
                
144                 .PotionBlue_Clic_Interval = 0
                    
                End If
        
146             .PotionRed_Clic = 0
148             .PotionRed_U = 0
150             .PotionRed_U_Interval = 0
152             .PotionRed_Clic_Interval = 0

            End If
        
            ' Pociones Azules (U)
154         If .PotionBlue_U > 0 Then
156             .PotionBlue_U_Interval = .PotionBlue_U_Interval + 1
            
158             If .PotionBlue_U_Interval >= TOLERANCE_POTIONBLUE_U Then
160                 .PotionBlue_U = .PotionBlue_U - 1
                
162                 If .PotionBlue_U > 0 Then
164                     .Stats.MinMan = .Stats.MinMan + Porcentaje(.Stats.MaxMan, 3) + .Stats.Elv \ 2 + 40 / .Stats.Elv
                                
166                     If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan
                    
168                     UpdateMAN = True

                    End If
                
170                 .PotionBlue_U_Interval = 0

                End If
            
172             .PotionRed_Clic = 0
174             .PotionRed_U = 0
176             .PotionRed_U_Interval = 0
178             .PotionRed_Clic_Interval = 0

            End If
        
180         If UpdateMAN Then Call WriteUpdateMana(UserIndex)
        
            ' Pociones Rojas (Clic)
182         If .PotionRed_Clic > 0 Then
184             .PotionRed_Clic_Interval = .PotionRed_Clic_Interval + 1
            
186             If .PotionRed_Clic_Interval >= TOLERANCE_POTIONRED_CLIC Then
188                 .PotionRed_Clic = .PotionRed_Clic - 1
                
190                 If .PotionRed_Clic > 0 Then
192                     .Stats.MinHp = .Stats.MinHp + ObjData(38).MaxModificador
    
194                     If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
196                     UpdateHP = True

                    End If
                
198                 .PotionRed_Clic_Interval = 0

                End If
            
200             .PotionBlue_Clic = 0
202             .PotionBlue_U = 0
204             .PotionBlue_U_Interval = 0
206             .PotionBlue_Clic_Interval = 0
        
            End If
        
            ' Pociones Rojas (U)
208         If .PotionRed_U > 0 Then
210             .PotionRed_U_Interval = .PotionRed_U_Interval + 1
            
212             If .PotionRed_U_Interval >= TOLERANCE_POTIONRED_U Then
214                 .PotionRed_U = .PotionRed_U - 1
                
216                 If .PotionRed_U > 0 Then
218                     .Stats.MinHp = .Stats.MinHp + ObjData(38).MaxModificador
        
220                     If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                    
222                     UpdateHP = True

                    End If
                
224                 .PotionRed_U_Interval = 0
                
226                 .PotionBlue_Clic = 0
228                 .PotionBlue_U = 0
230                 .PotionBlue_U_Interval = 0
232                 .PotionBlue_Clic_Interval = 0

                End If

            End If
        
234         If UpdateHP Then
236             Call WriteUpdateHP(UserIndex)
            
                If TempTick - .Counters.RuidoPocion > 1000 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                    .Counters.RuidoPocion = TempTick

                End If

            End If

        End With

        '<EhFooter>
        Exit Sub

DuracionPociones_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.DuracionPociones " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo HambreYSed_Err
        '</EhHeader>
  
100     With UserList(UserIndex)

102         If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        
            Dim cant As Byte
        
106         If .Stats.Elv <= 18 Then
108             cant = 1
            Else
110             cant = 10
            End If
        
            'Sed
112         If .Stats.MinAGU > 0 Then
114             If .Counters.AGUACounter < IntervaloSed Then
116                 .Counters.AGUACounter = .Counters.AGUACounter + 1
                Else
118                 .Counters.AGUACounter = 0
120                 .Stats.MinAGU = .Stats.MinAGU - cant
                
122                 If .Stats.MinAGU <= 0 Then
124                     .Stats.MinAGU = 0
126                     .flags.Sed = 1
                    End If
                
128                 fenviarAyS = True
                End If
            End If
        
            'hambre
130         If .Stats.MinHam > 0 Then
132             If .Counters.COMCounter < IntervaloHambre Then
134                 .Counters.COMCounter = .Counters.COMCounter + 1
                Else
136                 .Counters.COMCounter = 0
138                 .Stats.MinHam = .Stats.MinHam - cant

140                 If .Stats.MinHam <= 0 Then
142                     .Stats.MinHam = 0
144                     .flags.Hambre = 1
                    End If

146                 fenviarAyS = True
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

HambreYSed_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.HambreYSed " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Sanar(ByVal UserIndex As Integer, _
                 ByRef EnviarStats As Boolean, _
                 ByVal Intervalo As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo Sanar_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 2 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
            Dim mashit As Integer

            'con el paso del tiempo va sanando....pero muy lentamente ;-)
104         If .Stats.MinHp < .Stats.MaxHp Then
106             If .Counters.HPCounter < Intervalo Then
108                 .Counters.HPCounter = .Counters.HPCounter + 1
                Else
110                 mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                
112                 .Counters.HPCounter = 0
114                 .Stats.MinHp = .Stats.MinHp + mashit

116                 If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
118                 Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
120                 EnviarStats = True
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

Sanar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Sanar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub CargaNpcsDat()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CargaNpcsDat_Err
        '</EhHeader>

        Dim npcfile As String
    
100     npcfile = Npcs_FilePath
102     Set LeerNPCs = New clsIniManager
104     Call LeerNPCs.Initialize(npcfile)

106     Call DataServer_Generate_Npcs
        '<EhFooter>
        Exit Sub

CargaNpcsDat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CargaNpcsDat " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub CheckCountDown()
        '<EhHeader>
        On Error GoTo CheckCountDown_Err
        '</EhHeader>

100     If CountDown_Time = 0 Then Exit Sub
    
102     CountDown_Time = CountDown_Time - 1

            
104     If CountDown_Map > 0 Then
106         Call SendData(SendTarget.toMap, CountDown_Map, PrepareMessageRender_CountDown(CountDown_Time))

108         If CountDown_Time = 0 Then
110             Call SendData(SendTarget.toMap, CountDown_Map, PrepareMessageConsoleMsg("¡YA!", FontTypeNames.FONTTYPE_FIGHT))
            End If
        
        Else

112         Call SendData(SendTarget.toMapSecure, 0, PrepareMessageRender_CountDown(CountDown_Time))
        
114         If CountDown_Time = 0 Then
116             Call SendData(SendTarget.toMapSecure, 0, PrepareMessageConsoleMsg("¡YA!", FontTypeNames.FONTTYPE_FIGHT))
            End If
        End If

        '<EhFooter>
        Exit Sub

CheckCountDown_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CheckCountDown " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' # Chequea si la zona está caliente
Public Sub Map_CheckFire(ByVal Map As Integer)
    
    On Error GoTo ErrHandler
    
    Const MIN_USER_FIRE As Integer = 6
    
    With MapInfo(Map)
        If .Pk = False Then Exit Sub
        If .NumUsers < MIN_USER_FIRE Then Exit Sub
        
        .OnFire = .OnFire + 1
        
        If .OnFire >= 5 Then ' 3 minutos superando los 6 usuarios
            If FrasesLastMap = Map Then
                .OnFire = 0
                Exit Sub
            End If
            
            ' Selecciona una frase aleatoria
            Dim randomIndex As Integer
            randomIndex = Int((UBound(FrasesOnFire) + 1) * Rnd)
            Dim Mensaje As String
            Mensaje = Replace(FrasesOnFire(randomIndex), "{Mapa}", "**" & .Name & "**")
        
            ' # Envia mensaje a DISCORD de la concentración
            WriteMessageDiscord CHANNEL_ONFIRE, Mensaje & " " & "Players: " & .NumUsers
            
            .OnFire = 0
            FrasesLastMap = Map
        End If
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en checkfire")
End Sub
Sub PasarSegundo()
        '<EhHeader>
        On Error GoTo PasarSegundo_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i As Long
    
        
        ' Apertura del servidor
        Call User_Go_Initial_Version
        
        Call Teleports_Loop
        
100     Call Invations_MainLoop
    
        ' Respawn de Objetos
102     Call ChestLoop

        Call Pretorians_Loop
        
        ' Subasta de Objetos
104     Call Auction_Loop
    
        ' Cuenta regresiva
106     Call CheckCountDown
    
        ' Sistema de auto baneo
108     Call AutoBan_Loop
    
        ' Retos Loop
110     Call Retos_Loop
    
        ' Respawn de Npcs
112     Call Loop_RespawnNpc
    
114     If CountDownLimpieza > 0 Then
116         CountDownLimpieza = CountDownLimpieza - 1
        
118         If CountDownLimpieza < 4 And CountDownLimpieza > 0 Then
120             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpiando el mundo en " & CountDownLimpieza & " segundo" & IIf((CountDownLimpieza > 1), "s...", "..."), FontTypeNames.FONTTYPE_INFOGREEN))

            End If
        
122         If CountDownLimpieza <= 0 Then
124             Call LimpiarMundo

            End If

        End If
    
126     For i = 1 To LastUser

128         With UserList(i)

130             If i <> SLOT_TERMINAL_ARCHIVE Then
132                 UserList(i).Counters.TimeInactive = UserList(i).Counters.TimeInactive + 1
                
134                 If UserList(i).Counters.TimeInactive >= 60 Then
136                     UserList(i).Counters.TimeInactive = 0
138                     Call WriteDisconnect(i, True)
140                     Call Protocol.Kick(i)

                    End If

                End If
        
            End With
                
142         If UserList(i).flags.UserLogged Then

144             With UserList(i)
                    
                    If .Stats.BonusLast > 0 Then Call Reward_Check_User(i)  ' # Bonus del personaje
                    
146                 If .Counters.goHomeSec > 0 Then
                     
148                     .Counters.goHomeSec = .Counters.goHomeSec - 1
                        
150                     Call WriteUpdateGlobalCounter(i, 4, .Counters.goHomeSec)

                    End If

152                 If .Counters.ShieldBlocked > 0 Then
154                     .Counters.ShieldBlocked = .Counters.ShieldBlocked - 1

                    End If
                
156                 If .Counters.Shield > 0 Then
158                     .Counters.Shield = .Counters.Shield - 1
                    
160                     If .Counters.Shield = 0 Then
162                         Call RefreshCharStatus(i)

                        End If

                    End If
                
164                 If .Counters.FightSend > 0 Then
166                     .Counters.FightSend = .Counters.FightSend - 1
                    
168                     If .Counters.FightSend = 0 Then
170                         Call WriteConsoleMsg(i, "Ya puedes enviar otra invitación de reto.", FontTypeNames.FONTTYPE_INFOGREEN)

                        End If

                    End If
                
172                 If .Counters.ReviveAutomatic > 0 Then
174                     .Counters.ReviveAutomatic = .Counters.ReviveAutomatic - 1
                    
                        ' // NUEVO
176                     If .Counters.ReviveAutomatic > 0 And .Counters.ReviveAutomatic <= 5 Then
178                         Call WriteConsoleMsg(i, "Serás revivido en " & .Counters.ReviveAutomatic & " segundo" & IIf((.Counters.ReviveAutomatic = 1), "s.", "."), FontTypeNames.FONTTYPE_INFO)

                        End If
                    
180                     If .Counters.ReviveAutomatic = 0 Then
182                         If .flags.Muerto Then Call RevivirUsuario(i)

                        End If

                    End If
                
184                 If .Counters.FightInvitation > 0 Then
186                     .Counters.FightInvitation = .Counters.FightInvitation - 1

                    End If
                
188                 If .Counters.TimePublicationMao > 0 Then
190                     .Counters.TimePublicationMao = .Counters.TimePublicationMao - 1

                    End If
                
192                 If .Counters.TimeCreateChar > 0 Then .Counters.TimeCreateChar = .Counters.TimeCreateChar - 1
                
194                 If .Counters.TimeDenounce > 0 Then
196                     .Counters.TimeDenounce = .Counters.TimeDenounce - 1

                    End If
                
198                 Call Effect_Loop(i)
200                 Call AntiFrags_CheckTime(i)
                
202                 If .flags.Transform Then EfectoTransformacion (i)

204                 If .Counters.TimeFight > 0 Then
206                     .Counters.TimeFight = .Counters.TimeFight - 1
                    
                        ' Cuenta regresiva de retos y eventos
208                     If .Counters.TimeFight = 0 Then
210                         Call WriteRender_CountDown(i, .Counters.TimeFight)
                            'WriteConsoleMsg i, "Cuenta» ¡YA!", FontTypeNames.FONTTYPE_FIGHT
                                      
                            ' En los duelos desparalizamos el cliente
212                         If .flags.SlotEvent > 0 Then
214                             If Events(.flags.SlotEvent).Modality = eModalityEvent.Enfrentamientos Then
216                                 Call WriteUserInEvent(i)

                                End If

                            End If
                                      
218                         If .flags.SlotReto > 0 Then
220                             Call WriteUserInEvent(i)

                            End If

                        Else
222                         Call WriteRender_CountDown(i, .Counters.TimeFight)

                            'WriteConsoleMsg i, .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD
                        End If

                    End If

224                 If .Counters.TimeTransform > 0 Then
226                     .Counters.TimeTransform = .Counters.TimeTransform - 1

                    End If
                
228                 If .Counters.TimeGlobal > 0 Then
230                     .Counters.TimeGlobal = .Counters.TimeGlobal - 1

                    End If
        
232                 If .Counters.TimeBono > 0 Then
234                     .Counters.TimeBono = .Counters.TimeBono - 1
                    
                    End If
            
                    ' Tiempo para que el usuario se vaya del mapa
236                 If .Counters.TimeTelep > 0 Then

                        ' Efecto con objeto
238                     If .flags.ObjIndex Then
240                         If ObjData(.flags.ObjIndex).TelepMap = .Pos.Map Then
242                             .Counters.TimeTelep = .Counters.TimeTelep - 1
                                    
244                             If .Counters.TimeTelep = 0 Then
246                                 Call QuitarObjetos(.flags.ObjIndex, 1, i)
248                                 .flags.ObjIndex = 0
250                                 WarpUserChar i, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True
252                                 WriteConsoleMsg i, "El efecto de la teletransportación ha terminado.", FontTypeNames.FONTTYPE_INFO
    
                                End If

                            End If
                        
                        Else
254                         .Counters.TimeTelep = .Counters.TimeTelep - 1
                        
256                         If .Counters.TimeTelep = 0 Then
258                             If .flags.SlotEvent Then
260                                 Call Events_ChangePosition(i, .flags.SlotEvent)

                                End If

                            End If

                        End If

                    End If
                
                    ' Tiempo para cambiar de apariencia (cada 3 segundos en caso de evento)
262                 If .Counters.TimeApparience > 0 Then
264                     .Counters.TimeApparience = .Counters.TimeApparience - 1
                    
266                     If .Counters.TimeApparience = 0 Then
268                         Call Events_ChangeApparience(i)

                        End If

                    End If

                End With

            End If
        
270         With UserList(i)

                'Cerrar usuario
272             If .Counters.Saliendo Then
274                 .Counters.Salir = .Counters.Salir - 1

276                 If .Counters.Salir <= 0 Then
                                     
278                     If .flags.DeslogeandoCuenta Then
280                         .flags.DeslogeandoCuenta = False
                                Call WriteDisconnect(i, True)
                                Call FlushBuffer(i)
282                           Call Server.Kick(i, True)
                                           
                        Else
284                         Call WriteConsoleMsg(i, "Desconectado personaje...", FontTypeNames.FONTTYPE_INFO)
286                         Call WriteDisconnect(i)
290                         Call CloseSocket(i)
                              Call FlushBuffer(i)
                        End If

                    End If

                End If
        
            End With
      
292     Next i

        
        Call Streamer_CheckPosition
        
        '<EhFooter>
        Exit Sub

PasarSegundo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.PasarSegundo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
 
Public Function ReiniciarAutoUpdate() As Double
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ReiniciarAutoUpdate_Err
        '</EhHeader>

100     ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

        '<EhFooter>
        Exit Function

ReiniciarAutoUpdate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.ReiniciarAutoUpdate " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ReiniciarServidor_Err

        '</EhHeader>

        'commit experiencias
        Call mGroup.DistributeExpAndGldGroups
    
        'WorldSave
        Call ES.DoBackUp

106     If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

        'Chauuu
108     Unload frmMain

        '<EhFooter>
        Exit Sub

ReiniciarServidor_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.ReiniciarServidor " & "at line " & Erl

        

        '</EhFooter>
End Sub
 
Sub GuardarUsuarios(ByVal IsBackup As Boolean)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo GuardarUsuarios_Err

        '</EhHeader>
    
        Dim i             As Integer

        Dim UserGuardados As Long
        
104     For i = 1 To LastUser

106         If UserList(i).flags.UserLogged Then
                    If Not EsGm(i) Then
                        Call Power_Search(i)
                        
                    End If
                
                If GetTime - UserList(i).Counters.LastSave > IntervaloGuardarUsuarios Then
                    Call UpdatePremium(i)
                    
                    ' No guarda personajes en eventos.
                    If (UserList(i).flags.SlotEvent = 0 And UserList(i).flags.SlotReto = 0) Then
                        Call SaveUser(UserList(i), CharPath & UCase$(UserList(i).Name) & ".chr", False)
                    End If
                    
                    Call SaveDataAccount(i, UserList(i).Account.Email, UserList(i).IpAddress)
                End If
            End If

120     Next i
    
        '<EhFooter>
        Exit Sub

GuardarUsuarios_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.GuardarUsuarios " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub GuardarUsuarios_Close()

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo GuardarUsuarios_Err

        '</EhHeader>
    
        Dim i             As Integer

        Dim UserGuardados As Long
        
104     For i = 1 To LastUser

106         If UserList(i).flags.UserLogged Then
                Call UpdatePremium(i)
                    
                ' No guarda personajes en eventos.
                If (UserList(i).flags.SlotEvent = 0 And UserList(i).flags.SlotReto = 0) Then
                    Call SaveUser(UserList(i), CharPath & UCase$(UserList(i).Name) & ".chr", False)

                End If

            End If
            
            If UserList(i).AccountLogged Then
                Call SaveDataAccount(i, UserList(i).Account.Email, UserList(i).IpAddress)
            End If
            
120     Next i
    
        '<EhFooter>
        Exit Sub

GuardarUsuarios_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.GuardarUsuarios " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub FreeNPCs()
        '<EhHeader>
        On Error GoTo FreeNPCs_Err
        '</EhHeader>

        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all NPC Indexes
        '***************************************************
        Dim LoopC As Long
    
        ' Free all NPC indexes
100     For LoopC = 1 To MAXNPCS
102         Npclist(LoopC).flags.NPCActive = False
104     Next LoopC

        '<EhFooter>
        Exit Sub

FreeNPCs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.FreeNPCs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub FreeCharIndexes()
        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all char indexes
        '***************************************************
        ' Free all char indexes (set them all to 0)
        '<EhHeader>
        On Error GoTo FreeCharIndexes_Err
        '</EhHeader>
100     Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
        '<EhFooter>
        Exit Sub

FreeCharIndexes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.FreeCharIndexes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function Tilde(Data As String) As String
        '<EhHeader>
        On Error GoTo Tilde_Err
        '</EhHeader>
 
100     Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
        '<EhFooter>
        Exit Function

Tilde_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Tilde " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub WarpPosAnt(ByVal UserIndex As Integer)
    ' • Warpeo del personaje a su posición anterior.
          
    ' // NUEVO
    Dim Pos As WorldPos
          
    On Error GoTo WarpPosAnt_Error

    With UserList(UserIndex)
        Pos.Map = .PosAnt.Map
        Pos.X = .PosAnt.X
        Pos.Y = .PosAnt.Y
                          
        Call ClosestStablePos(Pos, Pos)
        Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
              
        .PosAnt.Map = 0
        .PosAnt.X = 0
        .PosAnt.Y = 0
          
    End With

    Exit Sub

WarpPosAnt_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure WarpPosAnt of Módulo General in line " & Erl
End Sub

Public Sub Transform_User(ByVal UserIndex As Integer, ByVal BodySelected As Integer)
        '<EhHeader>
        On Error GoTo Transform_User_Err
        '</EhHeader>
            
            
            Dim A As Long
            
100     With UserList(UserIndex)
        
102         If .flags.Transform = 0 Then
104             .CharMimetizado.Body = .Char.Body
106             .CharMimetizado.Head = .Char.Head
108             .CharMimetizado.CascoAnim = .Char.CascoAnim
110             .CharMimetizado.ShieldAnim = .Char.ShieldAnim
112             .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
114             .Char.Body = BodySelected
                '.Char.Head = 0
                '.Char.CascoAnim = 0
                '.Char.ShieldAnim = 0
                '.Char.WeaponAnim = 0
116             .flags.Transform = 1
                '.flags.Ignorado = True
            
            Else

118             If .flags.Navegando Then
120                 If .flags.Muerto = 0 Then
122                     Call ToggleBoatBody(UserIndex)
                    Else
124                     .Char.Body = iFragataFantasmal
126                     .Char.ShieldAnim = NingunEscudo
128                     .Char.WeaponAnim = NingunArma
130                     .Char.CascoAnim = NingunCasco

                          For A = 1 To MAX_AURAS
132                         .Char.AuraIndex(A) = NingunAura
                         Next A
                    End If

                Else
134                 .Char.Body = .CharMimetizado.Body
136                 .Char.Head = .CharMimetizado.Head
138                 .Char.CascoAnim = .CharMimetizado.CascoAnim
140                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
142                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim

                     For A = 1 To MAX_AURAS
144                     .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
                     Next A
                End If
            
146             .flags.Transform = 0
                '.flags.Ignorado = False
            
            End If
        
148         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
150         Call RefreshCharStatus(UserIndex)
        
        End With

        '<EhFooter>
        Exit Sub

Transform_User_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Transform_User " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub TransformVIP_User(ByVal UserIndex As Integer, ByVal BodySelected As Integer)
        '<EhHeader>
        On Error GoTo TransformVIP_User_Err
        '</EhHeader>
             
             
           Dim A As Long
100     With UserList(UserIndex)

102         If .flags.TransformVIP = 0 Then
104             .CharMimetizado.Body = .Char.Body
106             .CharMimetizado.Head = .Char.Head
108             .CharMimetizado.CascoAnim = .Char.CascoAnim
110             .CharMimetizado.ShieldAnim = .Char.ShieldAnim
112             .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
114             .Char.Body = BodySelected
116             .Char.Head = 0
118             .Char.CascoAnim = 0
120             .Char.ShieldAnim = 0
122             .Char.WeaponAnim = 0
124             .flags.TransformVIP = 1
                '.flags.Ignorado = True
                
                For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = NingunAura
                Next A
                          
            Else

126             If .flags.Navegando Then
128                 If .flags.Muerto = 0 Then
130                     Call ToggleBoatBody(UserIndex)
                    Else
132                     .Char.Body = iFragataFantasmal
134                     .Char.ShieldAnim = NingunEscudo
136                     .Char.WeaponAnim = NingunArma
138                     .Char.CascoAnim = NingunCasco

                          For A = 1 To MAX_AURAS
140                         .Char.AuraIndex(A) = NingunAura
                          Next A
                    End If

                Else
142                 .Char.Body = .CharMimetizado.Body
144                 .Char.Head = .CharMimetizado.Head
146                 .Char.CascoAnim = .CharMimetizado.CascoAnim
148                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
150                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim

                      For A = 1 To MAX_AURAS
152                     .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
                      Next A
                End If
            
              For A = 1 To MAX_AURAS
                         .Char.AuraIndex(A) = NingunAura
                          Next A
                          
154             .flags.TransformVIP = 0
                '.flags.Mimetizado = 0
                '.flags.Ignorado = False
            
            End If
        
156         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
158         Call RefreshCharStatus(UserIndex)
        
        End With

        '<EhFooter>
        Exit Sub

TransformVIP_User_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.TransformVIP_User " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Reinicio al deslogear del UserIndex
Public Sub AntiFrags_ResetInfo(ByRef IUser As User)
        '<EhHeader>
        On Error GoTo AntiFrags_ResetInfo_Err
        '</EhHeader>

        Dim A As Long
        Dim NullAntiFrag As tAntiFrags
    
100     For A = 1 To MAX_CONTROL_FRAGS
102         IUser.AntiFrags(A) = NullAntiFrag
104     Next A

        '<EhFooter>
        Exit Sub

AntiFrags_ResetInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.AntiFrags_ResetInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Buscamos al personaje asesinado en la lista.
Public Function AntiFrags_SlotRepeat(ByVal UserIndex As Integer, _
                                     ByVal VictimIndex As Integer) As Byte
        '<EhHeader>
        On Error GoTo AntiFrags_SlotRepeat_Err
        '</EhHeader>



        Dim A As Long
        Dim VictimName As String
    
    
100     VictimName = UCase$(UserList(VictimIndex).Name)
    
102     For A = 1 To MAX_CONTROL_FRAGS

104         With UserList(UserIndex).AntiFrags(A)

106             If .UserName = VictimName Then
108                 AntiFrags_SlotRepeat = A

                    Exit Function

                End If

            End With

110     Next A
    

        '<EhFooter>
        Exit Function

AntiFrags_SlotRepeat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.AntiFrags_SlotRepeat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function AntiFrags_SlotFree(ByVal UserIndex As Integer) As Byte
        '<EhHeader>
        On Error GoTo AntiFrags_SlotFree_Err
        '</EhHeader>


        Dim A As Long
    
100     For A = 1 To MAX_CONTROL_FRAGS

102         With UserList(UserIndex).AntiFrags(A)
            
104             If .Time <= 0 Then
106                 AntiFrags_SlotFree = A

                    Exit For

                End If

            End With

108     Next A

        '<EhFooter>
        Exit Function

AntiFrags_SlotFree_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.AntiFrags_SlotFree " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Un personaje es asesinado
Public Function AntiFrags_CheckUser(ByVal UserIndex As Integer, _
                                    ByVal VictimIndex As Integer, _
                                    ByVal Time As Long)
        '<EhHeader>
        On Error GoTo AntiFrags_CheckUser_Err
        '</EhHeader>

        Dim Slot As Integer
        Dim VictimName As String
100     VictimName = UCase$(UserList(VictimIndex).Name)
    
102     Slot = AntiFrags_SlotRepeat(UserIndex, VictimIndex)
    
104     If Slot <= 0 Then
106         Slot = AntiFrags_SlotFree(UserIndex)
          End If
    
108     If Slot <= 0 Then GoTo AntiFrags_CheckUser_Err
    
110     With UserList(UserIndex).AntiFrags(Slot)
        
            ' El personaje ya está en la lista por lo cual no cuenta el Frag.
112         If .UserName = UserList(VictimIndex).Name Then
114             AntiFrags_CheckUser = False

                Exit Function

            End If
        
            ' El personaje ya está en la lista por lo cual no cuenta el Frag.
116         If .Account = UserList(VictimIndex).Account.Email Then
118             AntiFrags_CheckUser = False

                Exit Function

            End If
        
            ' El personaje ya está en la lista por lo cual no cuenta el Frag.
120         If .IP <> vbNullString Then
122             If .IP = UserList(VictimIndex).IpAddress Then
124                 AntiFrags_CheckUser = False
    
                    Exit Function
    
                End If
            End If
        
126         .Time = Time
128         .UserName = UCase$(VictimName)
130         .Account = UserList(VictimIndex).Account.Email
132         .IP = UserList(VictimIndex).IpAddress
        
134         'Call WriteLogSecurity( "Victima con IP: " & .IP, UserList(UserIndex).Account.Email, .Account, eSubType_Security.eAntiFrags)
        End With
    
    
136     AntiFrags_CheckUser = True

        '<EhFooter>
        Exit Function

AntiFrags_CheckUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.AntiFrags_CheckUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Descontamos el tiempo del AntiFrags
Public Sub AntiFrags_CheckTime(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo AntiFrags_CheckTime_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_CONTROL_FRAGS

102         With UserList(UserIndex).AntiFrags(A)

104             If .Time > 0 Then
106                 .Time = .Time - 1
                
108                 If .Time <= 0 Then
110                     .Time = 0
112                     .UserName = vbNullString
                          .IP = vbNullString
                          .Account = vbNullString
                          .cant = 0
                    End If
                End If

            End With

114     Next A
    

        '<EhFooter>
        Exit Sub

AntiFrags_CheckTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.AntiFrags_CheckTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function CanUse_Inventory(ByVal UserIndex As Integer, _
                                 ByVal ObjIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo CanUse_Inventory_Err
        '</EhHeader>

100     If ObjIndex <= 0 Then Exit Function
    
102     With UserList(UserIndex)

104         Select Case ObjData(ObjIndex).OBJType
            
                Case eOBJType.otPergaminos
106                 CanUse_Inventory = (InvUsuario.ClasePuedeUsarItem(UserIndex, ObjIndex))
                        
108             Case eOBJType.otarmadura, eOBJType.otTransformVIP
110                 CanUse_Inventory = (InvUsuario.ClasePuedeUsarItem(UserIndex, ObjIndex) And InvUsuario.FaccionPuedeUsarItem(UserIndex, ObjIndex) And InvUsuario.SexoPuedeUsarItem(UserIndex, ObjIndex) And CheckRazaUsaRopa(UserIndex, ObjIndex))
                        
112             Case eOBJType.otcasco
114                 CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
116             Case eOBJType.otescudo
118                 CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
120             Case eOBJType.otWeapon
122                 CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
124             Case eOBJType.otAnillo, eOBJType.otMagic
126                 CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
128             Case eOBJType.otFlechas
130                 CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
   
132             Case Else
134                 CanUse_Inventory = True
            End Select
        
136         If ObjData(ObjIndex).Bronce = 1 And .flags.Bronce = 0 Then
138             CanUse_Inventory = False
            End If
        
140         If ObjData(ObjIndex).Plata = 1 And .flags.Plata = 0 Then
142             CanUse_Inventory = False
            End If
            
144         If ObjData(ObjIndex).Oro = 1 And .flags.Oro = 0 Then
146             CanUse_Inventory = False
            End If
        
148         If ObjData(ObjIndex).Premium = 1 And .flags.Premium = 0 Then
150             CanUse_Inventory = False
            End If
        End With

        '<EhFooter>
        Exit Function

CanUse_Inventory_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CanUse_Inventory " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' El anti pelotudos que no voy a necesitar porque el gm no necesitara sumonear y poner telep a futuro.
Public Function CanUserTelep(ByVal MapaActual As Integer, ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo CanUserTelep_Err
        '</EhHeader>

100     With UserList(UserIndex)
106         If .Counters.Pena > 0 Then Exit Function
108         If .flags.SlotReto > 0 Then Exit Function
110         If .flags.SlotEvent > 0 Then Exit Function
112         If .flags.SlotFast > 0 Then Exit Function
114         If .flags.Desafiando > 0 Then Exit Function
            If MapInfo(UserList(UserIndex).Pos.Map).Pk Then Exit Function
            If MapInfo(MapaActual).LvlMin > .Stats.Elv Then Exit Function
              
        End With
    
116     CanUserTelep = True
        '<EhFooter>
        Exit Function

CanUserTelep_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CanUserTelep " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub CheckingOcultation(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo CheckingOcultation_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

102         If .flags.Oculto > 0 Then
104             .flags.Oculto = 0
106             .Counters.TiempoOculto = 0
            
108             If .flags.Navegando = 0 Then
110                 If .flags.Invisible = 0 Then
112                     Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.charindex, False)
114                     Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If

        End With
    
        '<EhFooter>
        Exit Sub

CheckingOcultation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CheckingOcultation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function Faction_String(ByVal Faction As eFaction) As String
        '<EhHeader>
        On Error GoTo Faction_String_Err
        '</EhHeader>
    
100     Select Case Faction
    
            Case eFaction.fCrim
102             Faction_String = "CRIMINAL"

104         Case eFaction.fCiu
106             Faction_String = "CIUDADANO"

108         Case eFaction.fArmada
110             Faction_String = "ARMADA REAL"

112         Case eFaction.fLegion
114             Faction_String = "LEGION OSCURA"
        End Select
    
        '<EhFooter>
        Exit Function

Faction_String_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Faction_String " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function NpcInventory_GetAnimation(ByVal UserIndex As Integer, _
                                          ByVal ObjIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo NpcInventory_GetAnimation_Err
        '</EhHeader>

100     If ObjIndex = 0 Then Exit Function
    
102     With ObjData(ObjIndex)

104         Select Case .OBJType

                Case eOBJType.otarmadura
106                 If .RopajeEnano <> 0 And (UserList(UserIndex).Raza = eRaza.Enano Or UserList(UserIndex).Raza = eRaza.Gnomo) Then
108                     NpcInventory_GetAnimation = .RopajeEnano
                    Else
110                     NpcInventory_GetAnimation = .Ropaje
                    End If
112             Case eOBJType.otescudo
114                 NpcInventory_GetAnimation = .ShieldAnim
                
116             Case eOBJType.otcasco
118                 NpcInventory_GetAnimation = .CascoAnim
                
120             Case eOBJType.otWeapon

122                 If .WeaponRazaEnanaAnim <> 0 And _
                       (UserList(UserIndex).Raza = eRaza.Enano Or UserList(UserIndex).Raza = eRaza.Gnomo) Then
                    
124                     NpcInventory_GetAnimation = .WeaponRazaEnanaAnim
                    Else
126                     NpcInventory_GetAnimation = .WeaponAnim
                    End If

            End Select

        End With
    
        '<EhFooter>
        Exit Function

NpcInventory_GetAnimation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.NpcInventory_GetAnimation " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Respawn_Npc_Free(ByVal NpcIndex As Integer, _
                                 ByVal Map As Integer, _
                                 ByVal Time As Long, _
                                 ByVal CastleIndex As Integer, _
                                 ByRef OrigPos As WorldPos) As Boolean
        '<EhHeader>
        On Error GoTo Respawn_Npc_Free_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To RESPAWN_MAX

102         With Respawn_Npc(A)

104             If .Time = 0 Then
106                 .Map = Map
                    .OrigPos = OrigPos
108                 .NpcIndex = NpcIndex
110                 .Time = Time
                    .CastleIndex = CastleIndex
                    
112                 Respawn_Npc_Free = True

                    Exit Function

                End If

            End With

        Next
    
        '<EhFooter>
        Exit Function

Respawn_Npc_Free_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Respawn_Npc_Free " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Loop_RespawnNpc()

        '<EhHeader>
        On Error GoTo Loop_RespawnNpc_Err

        '</EhHeader>

        Dim A       As Long

        Dim OrigPos As WorldPos
    
100     For A = 1 To RESPAWN_MAX

102         With Respawn_Npc(A)

104             If .Time > 0 Then
106                 .Time = .Time - 1
                
108                 If .Time = 0 Then

                        Dim Npc As Integer

110                     Npc = CrearNPC(.NpcIndex, .Map, .OrigPos)
                    
112                     If Npc Then
                            Npclist(Npc).CastleIndex = .CastleIndex
                            
                            If Npclist(Npc).CastleIndex > 0 Then
                                Call mCastle.Castle_Close(Npclist(Npc).CastleIndex)
                            End If
                            
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡" & Npclist(Npc).Name & " en " & MapInfo(Npclist(Npc).Pos.Map).Name & "!", FontTypeNames.FONTTYPE_USERBRONCE))
                        
                            ' # Envia un mensaje a discord
                            Dim TextDiscord As String
                            TextDiscord = "--------------------"
                            TextDiscord = TextDiscord & vbCrLf & "¡**" & Npclist(Npc).Name & "** en **" & MapInfo(Npclist(Npc).Pos.Map).Name & "**!"
                            
                            If Npclist(Npc).NroDrops > 0 Or Npclist(Npc).Invent.NroItems > 0 Then
                                TextDiscord = TextDiscord & vbCrLf & Npclist(Npc).TempDrops
                            End If
                            
                            TextDiscord = TextDiscord & vbCrLf & "--------------------"
                            
                            WriteMessageDiscord CHANNEL_BOSSES, TextDiscord
                        End If
                    
116                     .NpcIndex = 0
118                     .Time = 0
120                     .Map = 0
                        .CastleIndex = 0
                        .OrigPos.Map = 0
                        .OrigPos.X = 0
                        .OrigPos.Y = 0
                    End If

                End If
        
            End With

122     Next A

        '<EhFooter>
        Exit Sub

Loop_RespawnNpc_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.General.Loop_RespawnNpc " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Function Is_Map_valid(ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Is_Map_valid_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If .Pos.Map >= 74 And .Pos.Map <= 87 Then Exit Function
104         If .Pos.Map = 24 Then Exit Function
        End With
    
106     Is_Map_valid = True
        '<EhFooter>
        Exit Function

Is_Map_valid_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Is_Map_valid " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckMailString_Err
        '</EhHeader>



        Dim lPos As Long

        Dim lX   As Long

        Dim iAsc As Integer
    
        '1er test: Busca un simbolo @
100     lPos = InStr(sString, "@")

102     If (lPos <> 0) Then

            '2do test: Busca un simbolo . después de @ + 1
104         If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
            '3er test: Recorre todos los caracteres y los valída
106         For lX = 0 To Len(sString) - 1

108             If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
110                 iAsc = Asc(mid$(sString, (lX + 1), 1))

112                 If Not CMSValidateChar_(iAsc) Then Exit Function
                End If

114         Next lX
        
            'Finale
116         CheckMailString = True
        End If


        '<EhFooter>
        Exit Function

CheckMailString_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CheckMailString " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
        '<EhHeader>
        On Error GoTo CMSValidateChar__Err
        '</EhHeader>
100     CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
        '<EhFooter>
        Exit Function

CMSValidateChar__Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CMSValidateChar_ " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Chequeamos si los personajes estan en un mapa determinado.
Public Sub Checking_UsersInMap(ByVal Map As Integer)
        '<EhHeader>
        On Error GoTo Checking_UsersInMap_Err
        '</EhHeader>

        Dim X As Long, Y As Long
    
100     For Y = YMinMapSize To YMaxMapSize
102         For X = XMinMapSize To XMaxMapSize

104             If MapData(Map, X, Y).UserIndex <> 0 Then
106                 Call WarpUserChar(MapData(Map, X, Y).UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, False)
                End If
            
108         Next X
110     Next Y

        '<EhFooter>
        Exit Sub

Checking_UsersInMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Checking_UsersInMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Elimina todos los objetos de un mapa que no se encuentren bloqueados.
Public Sub DeleteObjectMap(ByVal Map As Integer)
        '<EhHeader>
        On Error GoTo DeleteObjectMap_Err
        '</EhHeader>

        Dim LoopX As Long, LoopY As Long
       
100     For LoopX = XMinMapSize To XMaxMapSize
102         For LoopY = YMinMapSize To YMaxMapSize
    
104             If InMapBounds(Map, LoopX, LoopY) Then
                    
106                 If MapData(Map, LoopX, LoopY).Blocked = 0 Then
108                     EraseObj 10000, Map, LoopX, LoopY
                    End If
                    
                End If
    
110         Next LoopY
112     Next LoopX

        '<EhFooter>
        Exit Sub

DeleteObjectMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.DeleteObjectMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' # Creación de un objeto en un mapa determinado
Public Sub Create_ObjectMap(ByVal ObjIndex As Integer, _
                               ByVal Amount As Integer, _
                               ByVal Map As Integer, _
                               ByVal X As Byte, _
                               ByVal Y As Byte, _
                               ByVal ObjEvent As Byte)
        '<EhHeader>
        On Error GoTo Create_ObjectMap_Err
        '</EhHeader>

          
        Dim Pos As WorldPos

        Dim Obj As Obj
          
100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y
    
106     Obj.ObjIndex = ObjIndex
108     Obj.Amount = Amount
    
110     Call TirarItemAlPiso(Pos, Obj)
112     MapData(Pos.Map, Pos.X, Pos.Y).ObjEvent = ObjEvent
          

        '<EhFooter>
        Exit Sub

Create_ObjectMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Create_ObjectMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' # Quitamos las criaturas de un Mapa

Public Sub Remove_All_Map(ByVal Map As Integer, ByVal NpcIndex As Integer, ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo Remove_All_Map_Err
        '</EhHeader>

        Dim A As Long, B As Long
    
100     For A = YMinMapSize To YMaxMapSize
102         For B = XMinMapSize To XMaxMapSize

104             If InMapBounds(Map, A, B) Then
106                 If NpcIndex Then
108                     If MapData(Map, A, B).NpcIndex > 0 Then
110                         If Npclist(MapData(Map, A, B).NpcIndex).Attackable = 1 Then
112                             Call QuitarNPC(MapData(Map, A, B).NpcIndex)
                            End If
                        End If
                    End If
114                 If ObjIndex Then
116                     If MapData(Map, A, B).ObjInfo.ObjIndex > 0 Then
118                         If ItemNoEsDeMapa(Map, A, B, False) Then
120                             EraseObj 10000, Map, A, B
                            End If
                        End If
                    End If
                End If

122         Next B
124     Next A
          
        '<EhFooter>
        Exit Sub

Remove_All_Map_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Remove_All_Map " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Function GetTime() As Double
    On Error GoTo ErrHandler

    Dim CurrentTick As Double
    CurrentTick = timeGetTime() And &H7FFFFFFF
    
    If CurrentTick < LastTick Then
        overflowCount = overflowCount + 1
    End If
    
    LastTick = CurrentTick
    
    ' Time since last overflow plus overflows times MAX_TIME
    GetTime = CurrentTick + overflowCount * MAX_TIME
    Exit Function

ErrHandler:
    GetTime = 0
    Call LogError("E$rror gettime")
End Function

' # Mapas válidos para que los Game Master puedan hacer sus eventos y sumonear usuarios
Public Function EventMaster_CheckMapvalid(ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo EventMaster_CheckMapvalid_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
102         If (.Pos.Map = 74 Or _
                .Pos.Map = 76 Or _
                .Pos.Map = 77 Or _
                .Pos.Map = 79 Or _
                .Pos.Map = 80 Or _
                .Pos.Map = 81 Or _
                .Pos.Map = 82 Or _
                .Pos.Map = 83 Or _
                .Pos.Map = 84 Or _
                .Pos.Map = 87) Then
            
104             EventMaster_CheckMapvalid = True
                Exit Function
            
            End If
    
        End With
    
    
        '<EhFooter>
        Exit Function

EventMaster_CheckMapvalid_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.EventMaster_CheckMapvalid " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Determina si tiene algun item de los digitados
Public Function User_TieneObjetos_Especiales(ByVal UserIndex As Integer, _
                                             ByVal Bronce As Byte, _
                                             ByVal Plata As Byte, _
                                             ByVal Oro As Byte, _
                                             ByVal Premium) As Boolean
        '<EhHeader>
        On Error GoTo User_TieneObjetos_Especiales_Err
        '</EhHeader>
                                             

        Dim A     As Long
        Dim ObjIndex As Integer
    
        Dim Total As Long

100     For A = 1 To UserList(UserIndex).CurrentInventorySlots
102         ObjIndex = UserList(UserIndex).Invent.Object(A).ObjIndex
        
104         If ObjIndex > 0 Then
106             If Bronce = 0 And ObjData(ObjIndex).Bronce = 1 Then
108                 User_TieneObjetos_Especiales = True
                    Exit Function
                End If
                
110             If Plata = 0 And ObjData(ObjIndex).Plata = 1 Then
112                 User_TieneObjetos_Especiales = True
                    Exit Function
                End If
                
114             If Oro = 0 And ObjData(ObjIndex).Oro = 1 Then
116                 User_TieneObjetos_Especiales = True
                    Exit Function
                End If
            
118             If Premium = 0 And ObjData(ObjIndex).Premium = 1 Then
120                 User_TieneObjetos_Especiales = True
                    Exit Function
                End If
            End If


122     Next A
    
        '<EhFooter>
        Exit Function

User_TieneObjetos_Especiales_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.User_TieneObjetos_Especiales " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
' Transforma los segundos en un tiempo determinado en (Horas Minutos & Segundos)
Public Function SecondsToHMS(ByVal Seconds As Long) As String
        '<EhHeader>
        On Error GoTo SecondsToHMS_Err
        '</EhHeader>

        Dim HR As Integer
    
        Dim MS As Integer
    
        Dim SS As Integer
        
        Dim DS As Integer
        
        DS = (Seconds \ 3600) \ 24
        
100     HR = (Seconds \ 3600) Mod 24
    
102     MS = (Seconds Mod 3600) \ 60
    
104     SS = (Seconds Mod 3600) Mod 60

        
106     SecondsToHMS = IIf(DS > 0, DS & " días ", vbNullString) & IIf(HR > 0, HR & " horas ", vbNullString) & IIf(MS > 0, MS & " minutos ", vbNullString) & IIf(SS > 0, SS & " segundos", vbNullString)

        '<EhFooter>
        Exit Function

SecondsToHMS_Err:
        LogError Err.description & vbCrLf & _
               "in SecondsToHMS " & _
               "at line " & Erl

        '</EhFooter>
End Function

Public Sub CheckHappyHour()
        '<EhHeader>
        On Error GoTo CheckHappyHour_Err
        '</EhHeader>
100     HappyHour = Not HappyHour
    
102     If HappyHour Then
104         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡HappyHour Activado! Exp x2 ¡Entrená tu personaje!", FontTypeNames.FONTTYPE_USERBRONCE))
        Else
106         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡HappyHour Desactivado!", FontTypeNames.FONTTYPE_USERBRONCE))

        End If

        '<EhFooter>
        Exit Sub

CheckHappyHour_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CheckHappyHour " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub CheckPartyTime()
        '<EhHeader>
        On Error GoTo CheckPartyTime_Err
        '</EhHeader>
100     PartyTime = Not PartyTime
    
102     If PartyTime Then
104         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("PartyTime» Los miembros de la party reciben 25% de experiencia extra.", FontTypeNames.FONTTYPE_INVASION))
        Else
106         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("PartyTime» Desactivado!", FontTypeNames.FONTTYPE_INVASION))

        End If

        '<EhFooter>
        Exit Sub

CheckPartyTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.CheckPartyTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' @ Metodo Burbuja para ordenar arrays
Function BubbleSort(ByRef vIn As Variant, bAscending As Boolean, Optional vRet As Variant) As Boolean
        ' Sorts the single dimension list array, ascending or descending
        ' Returns sorted list in vRet if supplied, otherwise in vIn modified
        '<EhHeader>
        On Error GoTo BubbleSort_Err
        '</EhHeader>
        
        Dim First As Long, Last As Long
        Dim i As Long, j As Long, bWasMissing As Boolean
        Dim Temp As Variant, vW As Variant
    
100     First = LBound(vIn)
102     Last = UBound(vIn)
    
104     ReDim vW(First To Last, 1)
106     vW = vIn
    
108     If bAscending = True Then
110         For i = First To Last - 1
112             For j = i + 1 To Last
114                 If vW(i) > vW(j) Then
116                 Temp = vW(j)
118                 vW(j) = vW(i)
120                 vW(i) = Temp
                    End If
122             Next j
124         Next i
        Else 'descending sort
126         For i = First To Last - 1
128             For j = i + 1 To Last
130                 If vW(i) < vW(j) Then
132                 Temp = vW(j)
134                 vW(j) = vW(i)
136                 vW(i) = Temp
                    End If
138             Next j
140         Next i
        End If
  
       'find whether optional vRet was initially missing
142     bWasMissing = IsMissing(vRet)
   
       ' transfers
144    If bWasMissing Then
146      vIn = vW  'return in input array
       Else
148      ReDim vRet(First To Last, 1)
150      vRet = vW 'return with input unchanged
       End If
   
152    BubbleSort = True

        '<EhFooter>
        Exit Function

BubbleSort_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.BubbleSort " & _
               "at line " & Erl
               BubbleSort = False
        
        '</EhFooter>
End Function

Public Function Email_Is_Testing_Pro(ByVal Email As String) As Boolean
        '<EhHeader>
        On Error GoTo Email_Is_Testing_Pro_Err
        '</EhHeader>

100     If Email = "marinolauta@gmail.com" Or _
           Email = "montiel.marcoseze@gmail.com" Or _
           Email = "gabi.barrantes.94@gmail.com" Or _
           Email = "hogarcasa1991@gmail.com" Or _
           Email = "chontecito@gmail.com" Or _
           Email = "nuria_sabrina@hotmail.com" Or _
           Email = "dreamlotao@gmail.com" Then
       
102        Email_Is_Testing_Pro = True
        End If
        '<EhFooter>
        Exit Function

Email_Is_Testing_Pro_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.General.Email_Is_Testing_Pro " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function Detect_FirstDayNext()
    Dim currentDate As Date
    Dim newDate As Date
    Dim currentHour As Date
    
    ' Get the current date and time
    currentDate = Now
    
    ' Get the current hour and minute portion
    currentHour = TimeValue(currentDate)
    
    ' Add one month to the current date
    newDate = DateAdd("m", 1, currentDate)
    
    ' Set the day to 1
    'newDate = DateSerial(Year(newDate), Month(newDate), 1)
    
    ' Combine the new date with the current hour and minute
    newDate = DateValue(newDate) + currentHour
    
    ' Display the new date and time in a MsgBox (for verification)
    Detect_FirstDayNext = newDate
End Function
Function CalcularPorcentajeBonificación(Exp As Long, usuariosOnline As Long) As Double
    Dim Porcentaje As Double
    
    If usuariosOnline > 100 Then
        ' A partir de 100 onlines: 3% por cada online
        Porcentaje = (usuariosOnline * 0.3) / 100
    Else
        ' Antes de 100 onlines: 1.5% por cada online
        Porcentaje = (usuariosOnline * 0.15) / 100
    End If

    ' Aplicar la bonificación al valor de 'exp'
    Dim bonificación As Double
    bonificación = Exp * Porcentaje

    CalcularPorcentajeBonificación = bonificación
End Function

Function CalcularPorcentajeBonificacion(ByVal Exp As Long) As Double

    On Error GoTo ErrHandler
    Dim Porcentaje As Double
    
    Porcentaje = (NumUsers + UsersBot) * 0.002
    
    ' Aplicar la bonificación al valor de 'exp'
    Dim bonificación As Double
    bonificación = Exp * Porcentaje

    CalcularPorcentajeBonificacion = bonificación
    
    
    Exit Function
ErrHandler:
    
End Function

' # Cargamos las frases al azar
Public Sub CargarFrasesOnFire()
    Dim FilePath As String
    
    FilePath = DatPath & "frases_on_fire.txt"
    FrasesOnFire = LeerFrasesDesdeArchivo(FilePath)
End Sub


' Función para leer las frases desde un archivo y almacenarlas en un array
Private Function LeerFrasesDesdeArchivo(ByVal rutaArchivo As String) As String()
    Dim frases() As String
    Dim Contenido As String
    
    On Error Resume Next
    Open rutaArchivo For Binary As #1
    If Err.number = 0 Then
        Contenido = InputB(LOF(1), #1)
        Close #1
    Else
        ' Manejar el error, por ejemplo, mostrar un mensaje o registrar el error
        MsgBox "Error al leer el archivo de frases: " & Err.description, vbExclamation
        Exit Function
    End If
    On Error GoTo 0
    
    ' Convierte los bytes a una cadena Unicode
    Contenido = StrConv(Contenido, vbUnicode)
    
    ' Divide la cadena en frases utilizando vbCrLf
    frases = Split(Contenido, vbCrLf)
    
    LeerFrasesDesdeArchivo = frases
End Function




