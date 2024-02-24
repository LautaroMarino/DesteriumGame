Attribute VB_Name = "ProtocolArchive"
Option Explicit

' Protocolo encargado de enviar los datos a la TERMINAL encargada del manejo de archivos y/o manipulacion de bases de datos.

Public Const CHANNEL_CASTLE As String = "1183245947002687599"
Public Const CHANNEL_FIGHT As String = "1183246703684497418"
Public Const CHANNEL_LEVEL As String = "1183247950755934320"
Public Const CHANNEL_BOSSES As String = "1183248555197083678"
Public Const CHANNEL_ONFIRE As String = "1183298137985663026"
Public Const CHANNEL_TOURNAMENT As String = "1183317183921672232"
Public Const CHANNEL_PENAS As String = "1183603471514075196"
Public Const CHANNEL_MERCADER As String = "1183604814798663762"

Public Enum eSubType_Modality
    General = 1
    unoVSuno = 2
    dosVSdos = 3
    tresVStres = 4
    DagaRusa = 5
    DeathMatch = 6
    Imparable = 7
    ReyVsRey = 8
    Retos1vs1 = 9
    Fast1vs1 = 10
End Enum

Public Enum eSubType_Security

    eGeneral = 1
    eAntiCheat = 2
    eAntiFrags = 3
    eAntiFraude = 4

End Enum

Private Enum ServerPacketID

    LogSecurity = 1000
    UpdateStats = 1002
    SendEmail = 1003

    SendMap = 1004
    SendObj = 1005
    SendObj_Finish = 1006
    SendMercaderOffer = 1007
    SendUsersOn = 1008
    RequestConnect = 1009 ' El personaje solicita entrar a su cuenta
    UpdateUserData = 1010 ' El personaje solicita guardar su personaje
    
    RequestID = 1011
    UpdateUserEvent = 1012
    
    UpdateCastleEvent = 1013
    MessageDiscord = 1014
    AnalyzeText = 1015
End Enum

Public Enum CentralPacketID
    Identification = 1
    Identification_CanLogged = 2
    UpdateMapData = 3
    UpdateMercaderOffer = 4
    RequiredConnect = 5
    CloseConnection = 6               ' Cierra la conexión del cliente que está solicitando algo indebido
    SendConnectID = 7                 ' Envia el ID que solicito en un principio. Es el identificador que necesita el usuario para jugar.
    RequiredConnectBattle = 8           ' Conecta una nueva clase al Battle
    
    UpdateID = 9
    
    RewardData = 10
End Enum


' El 0 es el TDS, el 1 es el BattleServer.
Public Const ServerSelected As Integer = 0

Public Sub HandleCentralServer(ByVal Connection As Integer)
    On Error GoTo Errhandler
    
    Dim PacketID As Integer
    
    PacketID = Reader.ReadInt16
    
    
    Select Case PacketID
            
        Case CentralPacketID.Identification
            Call HandleIdentification(Connection)
            
        Case CentralPacketID.RequiredConnect
            Call HandleRequiredConnect(Connection)
        
        Case CentralPacketID.CloseConnection
            Call HandleCloseConnection(Connection)
            
        Case CentralPacketID.UpdateID
            Call HandleUpdateID(Connection)
            
        Case CentralPacketID.RewardData
            Call HandleRewardData(Connection)
            
    End Select
    
    Exit Sub
Errhandler:
    
End Sub

' # Recibe el ID de la cuenta
Private Sub HandleUpdateID(ByVal Connection As Integer)

    On Error GoTo Errhandler
    
    Dim Email As String
    Dim ID As Integer
    
    Email = Reader.ReadString8
    ID = Reader.ReadInt16
    
    Dim tUser As Integer
    
    tUser = CheckEmailLogged(LCase$(Email))
    
    If tUser > 0 Then
        UserList(tUser).ID = ID
    End If
    
    Exit Sub
Errhandler:
End Sub


' # Recibe la información de las recompensas a otorgar
Private Sub HandleRewardData(ByVal Connection As Integer)

    On Error GoTo Errhandler
    
    Dim ModalityID As Byte
    Dim PlayersCount As Integer
    Dim GamesWon As Integer
    Dim GamesPlayed As Integer
    Dim ConsecutiveWins As Integer
    
    Dim Player As tPlayerData
    Dim A As Long

    ModalityID = Reader.ReadInt8
    PlayersCount = Reader.ReadInt16
    
    For A = 1 To PlayersCount
        Player.PlayerName = Reader.ReadString8
        Player.GamesWon = Reader.ReadInt
        Player.GamesPlayed = Reader.ReadInt
        Player.ConsecutiveWins = Reader.ReadInt
        
        Call Reward_Process_User(ModalityID, Player)
        Call Log_Reward("Procesando a " & Player.PlayerName & " en " & ModalityID)
    Next A

    Exit Sub
Errhandler:
    
End Sub
' # Recibe la orden de desconectar al socket. Esto es cuando solicita información y no coordina con lo estipulado.
Private Sub HandleCloseConnection(ByVal Connection As Integer)
    
    On Error GoTo Errhandler
    Dim UserIndex As Integer
    
    UserIndex = Reader.ReadInt16
    
    Call Protocol.Kick(UserIndex, "Algo ocurrió con tu conexión. Si el error persiste, contacta al equipo de soporte.")
    
    Exit Sub
Errhandler:
End Sub

Private Sub HandleIdentification(ByVal Connection As Integer)
        '<EhHeader>
        On Error GoTo HandleIdentification_Err
        '</EhHeader>
        Dim Passwd As String
        Const Terminal_Server_Passwd As String = "TerminalServer$1983"
100     Passwd = Reader.ReadString8
    
102     If StrComp(Passwd, Terminal_Server_Passwd) = 0 Then
104         If SLOT_TERMINAL_ARCHIVE <> 0 Then
106             Call Protocol.Kick(SLOT_TERMINAL_ARCHIVE)
            Else
108             SLOT_TERMINAL_ARCHIVE = Connection

            End If
        
110         Call Logs_Security(eLog.eSecurity, eAntiHack, "Conexión de Terminal Server " & Now)
        End If
        '<EhFooter>
        Exit Sub

HandleIdentification_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleIdentification " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleRequiredConnect(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequiredConnect_Err
        '</EhHeader>
100     If SLOT_TERMINAL_ARCHIVE <> UserIndex Then Exit Sub
    
        Dim Email As String
        Dim Passwd As String
        Dim Key As String
        Dim tUser As Integer
        Dim FilePath As String
        
102     Email = Reader.ReadString8
104     Passwd = Reader.ReadString8
106     Key = Reader.ReadString8
108     Email = LCase$(Email)
        
110     FilePath = AccountPath & Email & ACCOUNT_FORMAT
    
        ' ¿No existe la Cuenta? » Creamos el archivo inicial con los datos
112     If Not FileExist(AccountPath & Email & ACCOUNT_FORMAT) Then
114         Call SaveDataNew(Email, Passwd, Key)
            Exit Sub
        End If
        
        ' La Cuenta Existe comprobamos que los datos no hayan sido actualizados
        Dim TempPasswd As String
116     TempPasswd = GetVar(FilePath, "INIT", "PASSWD")
    
        ' Datos incorrectos >> Actualizar
118     If Not StrComp(TempPasswd, Passwd) = 0 Then
        
            ' Si estaba online la persona y los datos cambiaron , la deslogeamos
120         tUser = CheckEmailLogged(Email)
        
122         If tUser > 0 Then
124             Call WriteErrorMsg(tUser, "Los datos de tu cuenta han sido actualizados. ¡Debes volver a ingresar!")
126             Call WriteDisconnect(tUser)
128             Call FlushBuffer(tUser)
130             Call CloseSocket(tUser)
        
            End If
        
132         Call SaveDataNew(Email, Passwd, Key)
        
        End If
        
        
        ' Agrego el ID esperando a que sea llamado
        
        '<EhFooter>
        Exit Sub

HandleRequiredConnect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequiredConnect " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteMessageDiscord(ByVal Channel As String, ByVal Text As String)
        On Error GoTo WriteMessageDiscord_Err
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteMessageDiscord. Terminal de Archivos no conectada.")
            Exit Sub
        End If
    
108     Call Writer.WriteInt(ServerPacketID.MessageDiscord)
110     Call Writer.WriteString8(Channel)
112     Call Writer.WriteString8(Text)
    
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        Exit Sub

WriteMessageDiscord_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteMessageDiscord " & _
               "at line " & Erl

End Sub
Public Sub WriteAnalyzeText(ByVal UserName As String, ByVal Text As String)
        On Error GoTo WriteAnalyzeText_Err
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteAnalyzeText. Terminal de Archivos no conectada.")
            Exit Sub
        End If
    
108     Call Writer.WriteInt(ServerPacketID.AnalyzeText)
110     Call Writer.WriteString8(UserName)
112     Call Writer.WriteString8(Text)
    
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        Exit Sub

WriteAnalyzeText_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteAnalyzeText " & _
               "at line " & Erl

End Sub
Public Sub WriteUpdateOnline()
        '<EhHeader>
        On Error GoTo WriteUpdateOnline_Err
        '</EhHeader>
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteUpdateOnline. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
        Dim Server As Byte
                  #If Classic = 1 Then
104                     Server = 0
                    #Else
106                     Server = 1
                    #End If
                
    
108     Call Writer.WriteInt(ServerPacketID.SendUsersOn)
110     Call Writer.WriteInt8(Server)
112     Call Writer.WriteInt16(NumUsers + UsersBot)
    
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateOnline_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteUpdateOnline " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteLogSecurity(ByVal Argument As String, _
                            ByVal Responsable As String, _
                            ByVal victim As String, _
                            ByVal SubType As eSubType_Security)
        '<EhHeader>
        On Error GoTo WriteLogSecurity_Err
        '</EhHeader>
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteLogSecurity. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
104     Call Writer.WriteInt(ServerPacketID.LogSecurity)
106     Call Writer.WriteString8(Argument)
108     Call Writer.WriteString8(Responsable)
110     Call Writer.WriteString8(victim)
112     Call Writer.WriteInt8(SubType)
    
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        '<EhFooter>
        Exit Sub

WriteLogSecurity_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteLogSecurity " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Sub WriteUpdateStats()
        '<EhHeader>
        On Error GoTo WriteUpdateStats_Err
        '</EhHeader>

100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteUpdateStats. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
104     Call Writer.WriteInt(ServerPacketID.UpdateStats)
    
106     Call Writer.WriteInt(NumUsers + UsersBot)
108     Call Writer.WriteInt(RECORDusuarios)
    
110     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)
    
        '<EhFooter>
        Exit Sub

WriteUpdateStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteUpdateStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteSendEmail(ByVal ID As TypeWorking, _
                          ByVal Email As String, _
                          ByVal Body As String)
        '<EhHeader>
        On Error GoTo WriteSendEmail_Err
        '</EhHeader>
                          
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteSendEmail. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
104     Call Writer.WriteInt(ServerPacketID.SendEmail)
    
106     Call Writer.WriteInt8(ID)
108     Call Writer.WriteString8(Email)
110     Call Writer.WriteString8(Body)
    
112     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)
    
        '<EhFooter>
        Exit Sub

WriteSendEmail_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteSendEmail " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteSendMercaderOffer(ByVal MercaderSlot As Integer, _
                                                     ByVal MercaderOffer As Byte, _
                                                     ByVal MercaderTime As Long)
        '<EhHeader>
        On Error GoTo WriteSendMercaderOffer_Err
        '</EhHeader>
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteSendMercaderOffer. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
104     Call Writer.WriteInt(ServerPacketID.SendMercaderOffer)
106     Call Writer.WriteInt16(MercaderSlot)
108     Call Writer.WriteInt8(MercaderOffer)
110     Call Writer.WriteInt32(MercaderTime)
    
112     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        '<EhFooter>
        Exit Sub

WriteSendMercaderOffer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteSendMercaderOffer " & _
               "at line " & Erl
        
    '</EhFooter>
End Sub

' # Actualiza los datos del usuario principales
Public Sub WriteUpdateUserData(ByRef IUser As User)
        '<EhHeader>
        On Error GoTo WriteUpdateUserData_Err
        '</EhHeader>
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteUpdateUserData. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
108     Call Writer.WriteInt(ServerPacketID.UpdateUserData)
110     Call Writer.WriteInt8(ServerSelected)

        With IUser
            Call Writer.WriteInt(.ID)  ' # PlayerID
            
            Call Writer.WriteInt8(.Clase) ' # Clase
            Call Writer.WriteInt8(.Raza) ' # Raza
            Call Writer.WriteInt8(.Stats.Elv) ' # Nivel
            Call Writer.WriteString8(.Name)     ' # Nombre
            Call Writer.WriteInt32(.Stats.Exp) ' # Experiencia
            
            Call Writer.WriteInt32(.Stats.Gld)           ' # Monedas de Oro
            Call Writer.WriteInt32(.Stats.Eldhir)    ' # Desterium Points
            Call Writer.WriteInt32(.Stats.Points)        ' # Puntos de Partida
            Call Writer.WriteInt32(.Faction.FragsOther)      ' # Usuarios Matados
            Call Writer.WriteInt32(.Faction.FragsCiu)      ' # Ciudadanos Matados
            Call Writer.WriteInt32(.Faction.FragsCri)      ' # Criminales Matados
            
            Dim Ups As Integer
            Ups = .Stats.MaxHp - getVidaIdeal(.Stats.Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion))
            
            Call Writer.WriteInt32(Ups)
            Call Writer.WriteInt16(.Stats.MaxHp)
            Call Writer.WriteInt16(.flags.Rachas)
            Call Writer.WriteInt16(.flags.RachasTemp)
            
            ' Inventory
            ' Spells
            ' Bank
            ' Anti Frags
            
            
            ' No lo se..
            ' Skills
            ' Atributos
            
            
        End With
        
    
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateUserData_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteUpdateUserData " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
' # Actualiza los datos del usuario principales
Public Sub WriteRequestID(ByVal Email As String)
        '<EhHeader>
        On Error GoTo WriteRequestID_Err
        '</EhHeader>
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteUpdateUserData. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
108     Call Writer.WriteInt(ServerPacketID.RequestID)
        Call Writer.WriteString8(Email)
        
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        '<EhFooter>
        Exit Sub

WriteRequestID_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteRequestID " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdateEvent(ByVal PlayerID As Integer, _
                            ByVal PlayerName As String, _
                            ByRef ModalityID As eSubType_Modality, _
                            ByVal Won As Boolean)
        '<EhHeader>
        On Error GoTo WriteUpdateStats_Err
        '</EhHeader>

100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteUpdateEvent. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
104     Call Writer.WriteInt(ServerPacketID.UpdateUserEvent)
        
        Call Writer.WriteInt8(ServerSelected)
        Call Writer.WriteInt16(PlayerID)
106     Call Writer.WriteString8(PlayerName)
108     Call Writer.WriteInt8(ModalityID)
        Call Writer.WriteBool(Won)
110     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)
    
        '<EhFooter>
        Exit Sub

WriteUpdateStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteUpdateStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdateCastleConquist(ByVal GuildName As String, ByVal Castillo As Byte, ByVal AddPoints As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateCastleConquist_Err
        '</EhHeader>
                            
100     If SLOT_TERMINAL_ARCHIVE = 0 Then
102         Call LogError("Error en WriteCastleConquist. Terminal de Archivos no conectada.")
            Exit Sub

        End If
    
        Dim Server As Byte
        
        #If Classic = 1 Then
104         Server = 0
        #Else
106         Server = 1
        #End If
                
    
108     Call Writer.WriteInt(ServerPacketID.UpdateCastleEvent)
110     Call Writer.WriteString8(GuildName)
112     Call Writer.WriteInt8(Castillo)
        Call Writer.WriteInt16(AddPoints)
        
114     Call SendData(ToOne, SLOT_TERMINAL_ARCHIVE, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateCastleConquist_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ProtocolArchive.WriteUpdateCastleConquist " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

