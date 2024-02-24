Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

Public RequestCount As Integer
Public LastRequestTime As Double

Public Type tImageData
    Bytes() As Byte
End Type

Public ImageData As tImageData

Public SLOT_TERMINAL_ARCHIVE As Integer ' Connection Index: Programa externo encargado de la manipulación de archivos.

Public Enum eSearchData
    eMac = 1
    eDisk = 2
    eIpAddress = 3
End Enum


Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long)

Private Const CLIENT_XOR_KEY As Long = 192
Private Const SERVER_XOR_KEY As Long = 128

Private Declare Function ntohl Lib "ws2_32" (ByVal netlong As Long) As Long

'We'll pass Long host address values in lieu of this struct:
Private Type in_addr

    s_b1 As Byte
    s_b2 As Byte
    s_b3 As Byte
    s_b4 As Byte

End Type

Private Declare Function RtlIpv4AddressToString _
                Lib "ntdll" _
                Alias "RtlIpv4AddressToStringW" (ByRef Addr As Any, _
                                                 ByVal pS As Long) As Long

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Public Const SEPARATOR As String * 1 = vbNullChar

Public Enum eMessageType

    Info = 0
    Admin = 1
    Guild = 2
    Party = 3
    Combate = 4
    Trabajo = 5
    m_MOTD = 6
    cEvents_Curso = 7
    cEvents_General = 8

End Enum

Private Enum ServerPacketID
    
    Connected
    loggedaccount
    LoggedAccountBatle
    AccountInfo
    logged                  ' LOGGED
    LoggedRemoveChar
    LoggedAccount_DataChar
    
    SendIntervals
    
    Mercader_List
    Mercader_ListOffer
    Mercader_ListInfo

    MiniMap_InfoCriature
    
    Render_CountDown
    
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateDsp               ' PACKETDSP
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterChangeHeading  ' CCH
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    CharacterAttackMovement
    CharacterAttackNpc
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMusic               ' TM
    PlayWave              ' TW
    StopWaveMap
    PauseToggle             ' BKW
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeBankSlot_Account
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    SetInvisible            ' NOVER
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV          '
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails
    
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    CancelOfferItem
    ShowMenu
    StrDextRunningOut
    ChatPersonalizado
    GroupPrincipal
    GroupUpdateExp
    
    UserInEvent
    SendInfoRetos
    
    MontateToggle
    SolicitaCapProc
    UpdateListSecurity
    CreateDamage
    
    ClickVesA
    
    UpdateControlPotas
    
    UpdateInfoIntervals
    UpdateGroupIndex

    ' Clanes
    Guild_List
    Guild_Info
    Guild_InfoUsers
    
    Fight_PanelAccept
    
    UpdateEffectPoison
    CreateFXMap
    RenderConsole
    ViewListQuest
    UpdateUserDead
    QuestInfo
    UpdateGlobalCounter
    SendInfoNpc
    UpdatePosGuild
    UpdateLevelGuild
    UpdateStatusMAO
    UpdateOnline
    UpdateEvento
    UpdateMeditation
    SendShopChars
    UpdateFinishQuest
    UpdateDataSkin
    RequiredMoveChar
    UpdateBar
    UpdateBarTerrain
    VelocidadToggle
    SpeedToChar
    UpdateUserTrabajo
    TournamentList
    
    StatsUser
    StatsUser_Inventory
    StatsUser_Spells
    StatsUser_Bank
    StatsUser_Skills
    StatsUser_Bonos
    StatsUser_Penas
    StatsUser_Skins
    StatsUser_Logros
    
    UpdateClient
End Enum



Public Enum ClientPacketID

    LoginAccount
    LoginChar
    LoginCharNew
    LoginRemove
    LoginName
    ChangeClass
   
    DragToggle
    RequestAtributes        'ATR
  
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    moveItem
    RightClick
    UserEditation
    
    
    ' Paquetes exclusivos de sistema de ROL (Desterium AO)
    PartyClient
    GroupChangePorc
    SendReply               'SendReply
    AcceptReply             'AcceptReply
    AbandonateReply         'AbandonateReply
    Entrardesafio
    SetPanelClient
    ChatGlobal
    LearnMeditation
    InfoEvento
    
    DragToPos
    Enlist
    Reward
    
    Fianza
    Home
    
    AbandonateFaction
    SendListSecurity
    BankDeposit             'DEPO
    MoveSpell               'DESPHE
    MoveBank
    UserCommerceOffer       'OFRECER
    Online                  '/ONLINE
    Quit                    '/SALIR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    PartyMessage            '/PMSG
    CouncilMessage          '/BMSG
    ChangeDescription       '/DESC
    Punishments             '/PENAS
    Gamble                  '/APOSTAR
    BankGold
    Denounce                '/DENUNCIAR
    Ping                    '/PING
    GmCommands
    InitCrafting
    ShareNpc                '/COMPARTIR
    StopSharingNpc
    Consultation
    Event_Participe
    
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseItem                 'USA
    UseItemTwo
    CraftBlacksmith         'CNS
    WorkLeftClick           'WLC
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    
    UpdateInactive
    
    Retos_RewardObj
    Mercader_New
    Mercader_Required
    
    Map_RequiredInfo
    Forgive_Faction
    WherePower
    
    Auction_New
    Auction_Info
    Auction_Offer
    
    GoInvation
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle

    Guilds_Required
    Guilds_Found
    Guilds_Invitation
    Guilds_Online
    Guilds_Kick
    Guilds_Abandonate
    Guilds_Talk
    
    Fight_CancelInvitation
    Events_DonateObject
    QuestRequired
    ModoStreamer
    StreamerSetLink
    ChangeNick
    ConfirmTransaccion
    ConfirmItem
    ConfirmTier
    RequiredShopChars
    ConfirmChar
    ConfirmQuest
    RequiredSkins
    RequiredLive
    AcelerationChar
    AlquilarComerciante
    TirarRuleta
    CastleInfo
    RequiredStatsUser
    CentralServer = 249
    [PacketCount]
End Enum

Public PacketUseItem As ClientPacketID
Public PacketWorkLeft As ClientPacketID

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFORED
    FONTTYPE_INFOGREEN
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_EVENT
    FONTTYPE_USERGOLD
    FONTTYPE_USERPREMIUM
    FONTTYPE_USERBRONCE
    FONTTYPE_USERPLATA
    FONTTYPE_ANGEL
    FONTTYPE_DEMONIO
    FONTTYPE_GLOBAL
    FONTTYPE_ADMIN
    FONTTYPE_CRITICO
    FONTTYPE_INFORETOS
    FONTTYPE_INVASION
    FONTTYPE_PODER
    FONTTYPE_DESAFIOS
    FONTTYPE_STREAM
    FONTTYPE_RMSG
End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss

End Enum

Public Server As Network.Server
Public Writer As Network.Writer
Public Reader As Network.Reader
Public Function IsRequestAllowed() As Boolean

    IsRequestAllowed = True
    Exit Function
    
    ' Configura el límite de solicitudes por segundo
    Const RequestLimitPerSecond As Integer = 10 ' Cambia este valor según tus necesidades

    ' Obtiene el tiempo actual en segundos con precisión de milisegundos
    Dim currentTime As Double
    currentTime = CDbl(Timer)

    ' Calcula el tiempo transcurrido desde la última solicitud
    Dim ElapsedSeconds As Double
    ElapsedSeconds = currentTime - LastRequestTime

    ' Si ha pasado más de 1 segundo, reinicia el contador de solicitudes
    If ElapsedSeconds >= 1 Then
        RequestCount = 0
        LastRequestTime = currentTime
    End If

    ' Verifica si se ha alcanzado el límite de solicitudes
    If RequestCount >= RequestLimitPerSecond Then
        ' La solicitud actual supera el límite
        IsRequestAllowed = False
    Else
        ' La solicitud está permitida
        RequestCount = RequestCount + 1
        IsRequestAllowed = True
    End If
End Function
Private Function verifyTimeStamp(ByVal ActualCount As Long, ByRef LastCount As Long, ByRef LastTick As Long, ByRef Iterations, ByVal UserIndex As Integer, ByVal PacketName As String, Optional ByVal DeltaThreshold As Long = 100, Optional ByVal MaxIterations As Long = 5, Optional ByVal CloseClient As Boolean = False) As Boolean
    
    Dim Ticks As Long, Delta As Long
    Ticks = GetTime
    
    Delta = (Ticks - LastTick)
    LastTick = Ticks

    'Controlamos secuencia para ver que no haya paquetes duplicados.
    If ActualCount <= LastCount Then
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Account.Email & " | Ip: " & UserList(UserIndex).IpAddress & ". ", FontTypeNames.FONTTYPE_INFOBOLD))
        Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Account.Email & " | Ip: " & UserList(UserIndex).IpAddress & ". ")
        LastCount = ActualCount
       ' Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    'controlamos speedhack/macro
    If Delta < DeltaThreshold Then
        Iterations = Iterations + 1
        If Iterations >= MaxIterations Then
            'Call WriteShowMessageBox(UserIndex, "Relajate andá a tomarte un té con Gulfas.")
            verifyTimeStamp = False
            'Call LogMacroServidor("El usuario " & UserList(UserIndex).name & " iteró el paquete " & PacketName & " " & MaxIterations & " veces.")
            Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Control de macro---> El usuario " & UserList(UserIndex).Name & "| Revisar --> " & PacketName & " (Envíos: " & Iterations & ").", FontTypeNames.FONTTYPE_INFOBOLD))
            Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "Control de macro---> El usuario " & UserList(UserIndex).Name & "| Revisar --> " & PacketName & " (Envíos: " & Iterations & ").")
            'Call WriteCerrarleCliente(UserIndex)
            'Call CloseSocket(UserIndex)
            LastCount = ActualCount
            Iterations = 0
            Debug.Print "CIERRO CLIENTE"
        End If
        'Exit Function
    Else
        Iterations = 0
    End If
        
    verifyTimeStamp = True
    LastCount = ActualCount
End Function
Public Sub Kick(ByVal Connection As Long, Optional ByVal Message As String = vbNullString)
        '<EhHeader>
        On Error GoTo Kick_Err
        '</EhHeader>
    
    
100     If (Message <> vbNullString) Then
102         Call WriteErrorMsg(Connection, Message)
            Call Server.Flush(Connection)
        End If
    
104     If UserList(Connection).flags.UserLogged Then
106         Call CloseSocket(Connection)
        End If
    
108     Call Server.Flush(Connection)
110     Call Server.Kick(Connection, True)
        '<EhFooter>
        Exit Sub

Kick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.Kick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Function Ipv4NetAtoS(ByVal NetAddrLong As Long) As String
On Error GoTo ErrHandler

    Dim pS   As Long

    Dim pEnd As Long

    Ipv4NetAtoS = Space$(15)
    pS = StrPtr(Ipv4NetAtoS)
    pEnd = RtlIpv4AddressToString(ntohl(NetAddrLong), pS)
    Ipv4NetAtoS = Left$(Ipv4NetAtoS, ((pEnd Xor &H80000000) - (pS Xor &H80000000)) \ 2)
    Exit Function
ErrHandler:
    Ipv4NetAtoS = "255.255.255.0"
End Function

Function NextOpenUser() As Integer
        
        On Error GoTo NextOpenUser_Err
        

        Dim LoopC As Long
   
100     For LoopC = 1 To MaxUsers + 1

102         If LoopC > MaxUsers Then Exit For
104         If (Not UserList(LoopC).ConnIDValida And UserList(LoopC).flags.UserLogged = False) Then Exit For
106     Next LoopC
   
108     NextOpenUser = LoopC

        
        Exit Function

NextOpenUser_Err:
        
End Function
Public Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)
        '<EhHeader>
        On Error GoTo OnServerConnect_Err
        '</EhHeader>

        Dim FreeUser As Long
            
100     If Connection <= MaxUsers Then
102         FreeUser = NextOpenUser()
            
104         UserList(FreeUser).ConnIDValida = True
106         UserList(FreeUser).IpAddress = Address

110         If FreeUser >= LastUser Then LastUser = FreeUser
                    
            Dim Server As Byte
            Server = 0

                    
112         Call WriteConnectedMessage(Connection, Server)
        Else
114         Call Protocol.Kick(Connection, "El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        End If
    
        '<EhFooter>
        Exit Sub

OnServerConnect_Err:
        Call Kick(Connection)
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.OnServerConnect " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Sub OnServerClose(ByVal Connection As Long)
On Error GoTo OnServerClose_Error

    
    If Not (Connection = SLOT_TERMINAL_ARCHIVE) Then
        If UserList(Connection).AccountLogged Then
            Call mAccount.DisconnectAccount(Connection)
        End If
    Else
        SLOT_TERMINAL_ARCHIVE = 0
    End If

    UserList(Connection).ConnIDValida = False
    UserList(Connection).IpAddress = vbNullString
    
    Call FreeSlot(Connection)
    
    
    Exit Sub

OnServerClose_Error:

    Call LogError("OnServerClose: " + Err.description)
    
End Sub

Public Sub OnServerSend(ByVal Connection As Long, ByVal Message As Network.Reader)
    '<EhHeader>
    On Error GoTo OnServerSend_Err
    '</EhHeader>


    '<EhFooter>
    Exit Sub

OnServerSend_Err:
    Call Kick(Connection)
    
    LogError Err.description & vbCrLf & _
           "in ServidorArgentum.Protocol.OnServerSend " & _
           "at line " & Erl

    '</EhFooter>
End Sub


Public Sub OnServerReceive(ByVal Connection As Long, ByVal Message As Network.Reader)
        '<EhHeader>
        On Error GoTo OnServerReceive_Err
        '</EhHeader>

        'Debug.Print "OnServerReceive"

      ' Dim BufferRef() As Byte
       ' Call message.GetData(BufferRef)
    
      '  Dim i As Long
       ' For i = 0 To UBound(BufferRef)
           ' BufferRef(i) = BufferRef(i) Xor CLIENT_XOR_KEY
       ' Next i
    
100  '   Set Reader = message
    
102  '   While (message.GetAvailable() > 0)

104         Call HandleIncomingData(Connection, Message)

        'Wend
    
106    ' Set Reader = Nothing
        '<EhFooter>
        Exit Sub

OnServerReceive_Err:
        Call Kick(Connection)
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.OnServerReceive " & _
               "at line " & Erl
        '</EhFooter>
End Sub

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer, ByVal Message As Network.Reader) As Boolean

        On Error Resume Next
        Set Reader = Message
        
        Dim PacketID As Long
    
100     PacketID = Reader.ReadInt
    
        Dim Time As Long
    
102     Time = GetTime()
        
104     If Time - UserList(UserIndex).Counters.TimeLastReset >= 5000 Then
106         UserList(UserIndex).Counters.TimeLastReset = Time
108         UserList(UserIndex).Counters.PacketCount = 0

        End If
    
110     UserList(UserIndex).Counters.PacketCount = UserList(UserIndex).Counters.PacketCount + 1
    
112     If UserList(UserIndex).Counters.PacketCount > 100 Then
            Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "Control de paquetes -> La cuenta " & UserList(UserIndex).Account.Email & " en personaje " & UserList(UserIndex).Name & " | IP: " & UserList(UserIndex).Account.Sec.IP_Address & " | Iteración paquetes | Último paquete: " & PacketID & ".")
114         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Control de paquetes -> La cuenta " & UserList(UserIndex).Account.Email & " en personaje " & UserList(UserIndex).Name & "  | Iteración paquetes | Último paquete: " & PacketID & ".", FontTypeNames.FONTTYPE_FIGHT))
116         UserList(UserIndex).Counters.PacketCount = 0
           '  Exit Function

        End If
    
118     If PacketID < 0 Or PacketID >= ClientPacketID.PacketCount Then
120         'Call Logs_Security(eSecurity, eAntiHack, "La cuenta " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " mando fake paquet " & PacketID)

122         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("La cuenta " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " mando fake paquet " & PacketID, FontTypeNames.FONTTYPE_SERVER))
            'Call Protocol.Kick(UserIndex)
            
           ' Exit Function

        End If
    
        'Does the packet requires a logged user??
126     If Not (PacketID = ClientPacketID.LoginChar Or PacketID = ClientPacketID.LoginCharNew Or PacketID = ClientPacketID.LoginName Or PacketID = ClientPacketID.LoginAccount Or PacketID = ClientPacketID.LoginRemove Or PacketID = ClientPacketID.CentralServer) Then

            ' Si no está logeado en la cuenta no se permite enviar paquetes
128         If Not UserList(UserIndex).AccountLogged Then
130             'Call Logs_Security(eSecurity, eAntiHack, "La IP: " & UserList(UserIndex).IpAddress & " mando fake paquet " & PacketID)
132             Call Protocol.Kick(UserIndex)
                Exit Function

            End If
    
134         If Not (PacketID = ClientPacketID.UpdateInactive Or PacketID = ClientPacketID.Mercader_Required Or PacketID = ClientPacketID.Mercader_New) Then
                
                'Is the user actually logged?
136             If Not UserList(UserIndex).flags.UserLogged Then
    
138                 Call CloseSocket(UserIndex)
                    Exit Function
    
                    'He is logged. Reset idle counter if id is valid.
140             ElseIf PacketID <= ClientPacketID.[PacketCount] Then
142                 UserList(UserIndex).Counters.IdleCount = 0

                End If

            End If
        
144     ElseIf PacketID <= ClientPacketID.[PacketCount] Then
146         UserList(UserIndex).Counters.IdleCount = 0

            'Is the user logged?
            
148         If UserList(UserIndex).flags.UserLogged Then
150             Call CloseSocket(UserIndex)

                Exit Function

            End If

        End If

        ' Ante cualquier paquete, pierde la proteccion de ser atacado.
152     UserList(UserIndex).flags.NoPuedeSerAtacado = False
    
154     Select Case PacketID
            
            Case ClientPacketID.RequiredStatsUser
                Call HandleRequiredStatsUser(UserIndex)
                
            Case ClientPacketID.CastleInfo
                Call HandleCastleInfo(UserIndex)
                
            Case ClientPacketID.TirarRuleta
                Call HandleTirarRuleta(UserIndex)
                
            Case ClientPacketID.AlquilarComerciante
                Call HandleAlquilarComerciante(UserIndex)
            
            Case ClientPacketID.AcelerationChar
                Call HandleAcelerationChar(UserIndex)
                
            Case ClientPacketID.RequiredLive
                Call HandleRequiredLive(UserIndex)
                
            Case ClientPacketID.RequiredSkins
                Call HandleRequiredSkin(UserIndex)
                
            Case ClientPacketID.ConfirmQuest
                Call HandleConfirmQuest(UserIndex)
                
            Case ClientPacketID.ConfirmChar
                Call HandleConfirmChar(UserIndex)
                
            Case ClientPacketID.RequiredShopChars
                Call HandleRequiredShopChars(UserIndex)
                
            Case ClientPacketID.ConfirmTier
                Call HandleConfirmTier(UserIndex)
                
            Case ClientPacketID.ConfirmItem
                Call HandleConfirmItem(UserIndex)
                
            Case ClientPacketID.ConfirmTransaccion
                Call HandleConfirmTransaccion(UserIndex)
                
            Case ClientPacketID.ChangeNick
                Call HandleChangeNick(UserIndex)
                
            Case ClientPacketID.StreamerSetLink
                Call HandleStreamerSetLink(UserIndex)
                
            Case ClientPacketID.ModoStreamer
                Call HandleModoStreamer(UserIndex)
                
            Case ClientPacketID.ChangeClass
                Call HandleChangeClass(UserIndex)
                    
            Case ClientPacketID.LoginName
                Call HandleLoginName(UserIndex)
                
            Case ClientPacketID.CentralServer
                Call HandleCentralServer(UserIndex)

158         Case ClientPacketID.Fight_CancelInvitation
160             Call HandleFight_CancelInvitation(UserIndex)
            
162         Case ClientPacketID.Guilds_Talk
164             Call HandleGuilds_Talk(UserIndex)
            
166         Case ClientPacketID.Guilds_Abandonate
168             Call HandleGuilds_Abandonate(UserIndex)
            
170         Case ClientPacketID.Guilds_Kick
172             Call HandleGuilds_Kick(UserIndex)
            
174         Case ClientPacketID.Guilds_Online
176             Call HandleGuilds_Online(UserIndex)
        
178         Case ClientPacketID.Guilds_Invitation
180             Call HandleGuilds_Invitation(UserIndex)
            
182         Case ClientPacketID.Guilds_Found
184             Call HandleGuilds_Found(UserIndex)
            
186         Case ClientPacketID.Guilds_Required
188             Call HandleGuilds_Required(UserIndex)
            
190         Case ClientPacketID.Retos_RewardObj
192             Call HandleRetos_RewardObj(UserIndex)
            
194         Case ClientPacketID.UpdateInactive
196             Call HandleUpdateInactive(UserIndex)
            
198         Case ClientPacketID.SendReply
200             Call HandleSendReply(UserIndex)
            
202         Case ClientPacketID.AcceptReply
204             Call HandleAcceptReply(UserIndex)
            
206         Case ClientPacketID.AbandonateReply
208             Call HandleAbandonateReply(UserIndex)
        
210         Case ClientPacketID.SendListSecurity
212             Call HandleSendListSecurity(UserIndex)
            
214         Case ClientPacketID.Event_Participe
216             Call HandleEvent_Participe(UserIndex)
            
218         Case ClientPacketID.AbandonateFaction
220             Call HandleAbandonateFaction(UserIndex)

222         Case ClientPacketID.LoginRemove
224             Call HandleLoginRemove(UserIndex)
            
226         Case ClientPacketID.LoginAccount
228             Call HandleLoginAccount(UserIndex)
        
230         Case ClientPacketID.Mercader_New
232             Call HandleMercader_New(UserIndex)
            
238         Case ClientPacketID.Mercader_Required
240             Call HandleMercader_Required(UserIndex)
            
242         Case ClientPacketID.Forgive_Faction
244             Call HandleForgive_Faction(UserIndex)
        
246         Case ClientPacketID.WherePower
248             Call HandleWherePower(UserIndex)
        
250         Case ClientPacketID.Auction_New
252             Call HandleAuction_New(UserIndex)
        
254         Case ClientPacketID.Auction_Info
256             Call HandleAuction_Info(UserIndex)
            
258         Case ClientPacketID.Auction_Offer
260             Call HandleAuction_Offer(UserIndex)
            
262         Case ClientPacketID.GoInvation
264             Call HandleGoInvation(UserIndex)
            
266         Case ClientPacketID.Map_RequiredInfo
268             Call HandleMap_RequiredInfo(UserIndex)
            
282         Case ClientPacketID.LoginChar
284             Call HandleLoginChar(UserIndex)
            
286         Case ClientPacketID.LoginCharNew
288             Call HandleLoginCharNew(UserIndex)
            
290         Case ClientPacketID.Entrardesafio
292             Call HandleEntrarDesafio(UserIndex)
            
294         Case ClientPacketID.SetPanelClient
296             Call HandleSetPanelClient(UserIndex)
            
298         Case ClientPacketID.GroupChangePorc
300             Call HandleGroupChangePorc(UserIndex)
            
302         Case ClientPacketID.PartyClient
304             Call HandlePartyClient(UserIndex)
        
306         Case ClientPacketID.Talk                    ';
308             Call HandleTalk(UserIndex)
        
310         Case ClientPacketID.Yell                    '-
312             Call HandleYell(UserIndex)
        
314         Case ClientPacketID.Whisper                 '\
316             Call HandleWhisper(UserIndex)
        
318         Case ClientPacketID.Walk                    'M
320             Call HandleWalk(UserIndex)
        
322         Case ClientPacketID.RequestPositionUpdate   'RPU
324             Call HandleRequestPositionUpdate(UserIndex)
        
326         Case ClientPacketID.Attack                  'AT
328             Call HandleAttack(UserIndex)
        
330         Case ClientPacketID.PickUp                  'AG
332             Call HandlePickUp(UserIndex)
        
334         Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
336             Call HandleSafeToggle(UserIndex)
        
338         Case ClientPacketID.ResuscitationSafeToggle
340             Call HandleResuscitationToggle(UserIndex)
            
342         Case ClientPacketID.DragToggle
344             Call HandleDragToggle(UserIndex)
        
346         Case ClientPacketID.RequestAtributes        'ATR
348             Call HandleRequestAtributes(UserIndex)
        
350         Case ClientPacketID.RequestSkills           'ESKI
352             Call HandleRequestSkills(UserIndex)
        
354         Case ClientPacketID.RequestMiniStats        'FEST
356             Call HandleRequestMiniStats(UserIndex)
        
358         Case ClientPacketID.CommerceEnd             'FINCOM
360             Call HandleCommerceEnd(UserIndex)
            
362         Case ClientPacketID.CommerceChat
364             Call HandleCommerceChat(UserIndex)
        
366         Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
368             Call HandleUserCommerceEnd(UserIndex)
            
370         Case ClientPacketID.UserCommerceConfirm
372             Call HandleUserCommerceConfirm(UserIndex)
        
374         Case ClientPacketID.BankEnd                 'FINBAN
376             Call HandleBankEnd(UserIndex)
        
378         Case ClientPacketID.UserCommerceOk          'COMUSUOK
380             Call HandleUserCommerceOk(UserIndex)
        
382         Case ClientPacketID.UserCommerceReject      'COMUSUNO
384             Call HandleUserCommerceReject(UserIndex)
        
386         Case ClientPacketID.Drop                    'TI
388             Call HandleDrop(UserIndex)
        
390         Case ClientPacketID.CastSpell               'LH
392             Call HandleCastSpell(UserIndex)
        
394         Case ClientPacketID.LeftClick               'LC
396             Call HandleLeftClick(UserIndex)
        
398         Case ClientPacketID.DoubleClick             'RC
400             Call HandleDoubleClick(UserIndex)
        
402         Case ClientPacketID.Work                    'UK
404             Call HandleWork(UserIndex)
        
406         Case ClientPacketID.UseItem                 'USA
408             Call HandleUseItem(UserIndex)
        
            Case ClientPacketID.UseItemTwo                 'USA
                Call HandleUseItemTwo(UserIndex)

410         Case ClientPacketID.CraftBlacksmith         'CNS
412             Call HandleCraftBlacksmith(UserIndex)
        
414         Case ClientPacketID.WorkLeftClick           'WLC
416             Call HandleWorkLeftClick(UserIndex)
        
418         Case ClientPacketID.SpellInfo               'INFS
420             Call HandleSpellInfo(UserIndex)
        
422         Case ClientPacketID.EquipItem               'EQUI
424             Call HandleEquipItem(UserIndex)
        
426         Case ClientPacketID.ChangeHeading           'CHEA
428             Call HandleChangeHeading(UserIndex)
        
430         Case ClientPacketID.ModifySkills            'SKSE
432             Call HandleModifySkills(UserIndex)
        
434         Case ClientPacketID.Train                   'ENTR
436             Call HandleTrain(UserIndex)
        
438         Case ClientPacketID.CommerceBuy             'COMP
440             Call HandleCommerceBuy(UserIndex)
        
442         Case ClientPacketID.BankExtractItem         'RETI
444             Call HandleBankExtractItem(UserIndex)
        
446         Case ClientPacketID.CommerceSell            'VEND
448             Call HandleCommerceSell(UserIndex)
        
450         Case ClientPacketID.BankDeposit             'DEPO
452             Call HandleBankDeposit(UserIndex)
        
454         Case ClientPacketID.MoveSpell               'DESPHE
456             Call HandleMoveSpell(UserIndex)
            
458         Case ClientPacketID.MoveBank
460             Call HandleMoveBank(UserIndex)
        
462         Case ClientPacketID.UserCommerceOffer       'OFRECER
464             Call HandleUserCommerceOffer(UserIndex)
         
466         Case ClientPacketID.Online                  '/ONLINE
468             Call HandleOnline(UserIndex)
        
470         Case ClientPacketID.Quit                    '/SALIR
472             Call HandleQuit(UserIndex)
        
474         Case ClientPacketID.Meditate                '/MEDITAR
476             Call HandleMeditate(UserIndex)
        
478         Case ClientPacketID.Resucitate              '/RESUCITAR
480             Call HandleResucitate(UserIndex)
        
482         Case ClientPacketID.Heal                    '/CURAR
484             Call HandleHeal(UserIndex)
        
486         Case ClientPacketID.Help                    '/AYUDA
488             Call HandleHelp(UserIndex)
        
490         Case ClientPacketID.RequestStats            '/EST
492             Call HandleRequestStats(UserIndex)
        
494         Case ClientPacketID.CommerceStart           '/COMERCIAR
496             Call HandleCommerceStart(UserIndex)
        
498         Case ClientPacketID.BankStart               '/BOVEDA
500             Call HandleBankStart(UserIndex)
        
502         Case ClientPacketID.PartyMessage            '/PMSG
504             Call HandlePartyMessage(UserIndex)
        
506         Case ClientPacketID.CouncilMessage          '/BMSG
508             Call HandleCouncilMessage(UserIndex)
        
510         Case ClientPacketID.ChangeDescription       '/DESC
512             Call HandleChangeDescription(UserIndex)
        
514         Case ClientPacketID.Punishments             '/PENAS
516             Call HandlePunishments(UserIndex)
        
518         Case ClientPacketID.Gamble                  '/APOSTAR
520             Call HandleGamble(UserIndex)
        
522         Case ClientPacketID.BankGold
524             Call HandleBankGold(UserIndex)
            
526         Case ClientPacketID.Denounce                '/DENUNCIAR
528             Call HandleDenounce(UserIndex)
        
530         Case ClientPacketID.Ping                    '/PING
532             Call HandlePing(UserIndex)
        
534         Case ClientPacketID.GmCommands              'GM Messages
536             Call HandleGMCommands(UserIndex)
            
538         Case ClientPacketID.InitCrafting
540             Call HandleInitCrafting(UserIndex)
            
542         Case ClientPacketID.ShareNpc                '/COMPARTIR
544             Call HandleShareNpc(UserIndex)
            
546         Case ClientPacketID.StopSharingNpc
548             Call HandleStopSharingNpc(UserIndex)
            
550         Case ClientPacketID.Consultation
552             Call HandleConsultation(UserIndex)
        
554         Case ClientPacketID.moveItem
556             Call HandleMoveItem(UserIndex)
            
558         Case ClientPacketID.RightClick
560             Call HandleRightClick(UserIndex)
            
562         Case ClientPacketID.UserEditation
564             Call HandleUserEditation(UserIndex)
            
566         Case ClientPacketID.ChatGlobal
568             Call HandleChatGlobal(UserIndex)
            
570         Case ClientPacketID.LearnMeditation
572             Call HandleLearnMeditation(UserIndex)
            
574         Case ClientPacketID.InfoEvento
576             Call HandleInfoEvento(UserIndex)
        
578         Case ClientPacketID.DragToPos
580             Call HandleDragToPos(UserIndex)
            
582         Case ClientPacketID.Enlist
584             Call HandleEnlist(UserIndex)
            
586         Case ClientPacketID.Reward
588             Call HandleReward(UserIndex)
            
590         Case ClientPacketID.Fianza
592             Call HandleFianza(UserIndex)
            
594         Case ClientPacketID.Home
596             'Call HandleHome(UserIndex)
        
598         Case ClientPacketID.Events_DonateObject
600             Call HandleEvents_DonateObject(UserIndex)
            
602         Case ClientPacketID.QuestRequired
604             Call HandleQuestRequired(UserIndex)
                 
                Case Else
                    Err.Raise -1, "Invalid Message"

        End Select
    
        
    If (Message.GetAvailable() > 0) Then
      '  Err.Raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketID & "' se encuentra en mal estado con '" & message.GetAvailable() & "' bytes de mas por el usuario '" & UserList(UserIndex).Name & "'"
    End If
    
HandleIncomingData_Err:
    
    Set Reader = Nothing

    If Err.number <> 0 Then
        Call LogError(Err.description & vbNewLine & "PackedID: " & PacketID & vbNewLine & IIf(UserList(UserIndex).flags.UserLogged, "Usuario: " & UserList(UserIndex).Name, "UserIndex: " & UserIndex & " con IP: " & UserList(UserIndex).IpAddress & " Email: " & UserList(UserIndex).Account.Email))
        'Call CloseSocket(UserIndex)
        
        HandleIncomingData = False
    End If
End Function

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, _
                             ByVal MessageIndex As Integer, _
                             Optional ByVal Arg1 As Long, _
                             Optional ByVal Arg2 As Long, _
                             Optional ByVal Arg3 As Long, _
                             Optional ByVal StringArg1 As String)
        '<EhHeader>
        On Error GoTo WriteMultiMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.MultiMessage)
102     Call Writer.WriteInt(MessageIndex)
        
104     Select Case MessageIndex

            Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
106         Case eMessages.NPCHitUser
108             Call Writer.WriteInt(Arg1) 'Target
110             Call Writer.WriteInt(Arg2) 'damage
                
112         Case eMessages.UserHitNPC
114             Call Writer.WriteInt(Arg1) 'damage
                
116         Case eMessages.UserAttackedSwing
118             Call Writer.WriteInt(UserList(Arg1).Char.charindex)
                
120         Case eMessages.UserHittedByUser
122             Call Writer.WriteInt(Arg1) 'AttackerIndex
124             Call Writer.WriteInt(Arg2) 'Target
126             Call Writer.WriteInt(Arg3) 'damage
                
128         Case eMessages.UserHittedUser
130             Call Writer.WriteInt(Arg1) 'AttackerIndex
132             Call Writer.WriteInt(Arg2) 'Target
134             Call Writer.WriteInt(Arg3) 'damage
                
136         Case eMessages.WorkRequestTarget
138             Call Writer.WriteInt(Arg1) 'skill
            
140         Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
142             Call Writer.WriteInt(UserList(Arg1).Char.charindex) 'VictimIndex
144             Call Writer.WriteInt(Arg2) 'Expe
            
146         Case eMessages.UserKill '"¡" & .name & " te ha matado!"
148             Call Writer.WriteInt(UserList(Arg1).Char.charindex) 'AttackerIndex
            
150         Case eMessages.Home
152             Call Writer.WriteInt(CByte(Arg1))
154             Call Writer.WriteInt(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda así.
156             Call Writer.WriteString8(StringArg1) 'Call .Writer.WriteInt(CByte(Arg2))
        
158         Case eMessages.EarnExp
        End Select

160     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteMultiMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMultiMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGMCommands_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim Command As Byte

100     With UserList(UserIndex)
    
102         Command = Reader.ReadInt
    
104         Select Case Command
                
                Case eGMCommands.GMMessage                '/GMSG
106                 Call HandleGMMessage(UserIndex)
        
108             Case eGMCommands.ShowName                '/SHOWNAME
110                 Call HandleShowName(UserIndex)
        
112             Case eGMCommands.serverTime              '/HORA
114                 Call HandleServerTime(UserIndex)
        
116             Case eGMCommands.Where                   '/DONDE
118                 Call HandleWhere(UserIndex)
        
120             Case eGMCommands.CreaturesInMap          '/NENE
122                 Call HandleCreaturesInMap(UserIndex)
        
124             Case eGMCommands.WarpChar                '/TELEP
126                 Call HandleWarpChar(UserIndex)
        
128             Case eGMCommands.Silence                 '/SILENCIAR
130                 Call HandleSilence(UserIndex)
        
132             Case eGMCommands.GoToChar                '/IRA
134                 Call HandleGoToChar(UserIndex)
        
136             Case eGMCommands.Invisible               '/INVISIBLE
138                 Call HandleInvisible(UserIndex)
        
140             Case eGMCommands.GMPanel                 '/PANELGM
142                 Call HandleGMPanel(UserIndex)
        
144             Case eGMCommands.RequestUserList         'LISTUSU
146                 Call HandleRequestUserList(UserIndex)
        
148             Case eGMCommands.Jail                    '/CARCEL
150                 Call HandleJail(UserIndex)
        
152             Case eGMCommands.KillNPC                 '/RMATA
154                 Call HandleKillNPC(UserIndex)
        
156             Case eGMCommands.WarnUser                '/ADVERTENCIA
158                 Call HandleWarnUser(UserIndex)
        
160             Case eGMCommands.RequestCharInfo         '/INFO
162                 Call HandleRequestCharInfo(UserIndex)
        
172             Case eGMCommands.RequestCharInventory    '/INV
174                 Call HandleRequestCharInventory(UserIndex)
        
176             Case eGMCommands.RequestCharBank         '/BOV
178                 Call HandleRequestCharBank(UserIndex)
        
184             Case eGMCommands.ReviveChar              '/REVIVIR
186                 Call HandleReviveChar(UserIndex)
        
188             Case eGMCommands.OnlineGM                '/ONLINEGM
190                 Call HandleOnlineGM(UserIndex)
        
192             Case eGMCommands.OnlineMap               '/ONLINEMAP
194                 Call HandleOnlineMap(UserIndex)
        
196             Case eGMCommands.Forgive                 '/PERDON
198                 Call HandleForgive(UserIndex)
        
200             Case eGMCommands.Kick                    '/ECHAR
202                 Call HandleKick(UserIndex)
        
204             Case eGMCommands.Execute                 '/EJECUTAR
206                 Call HandleExecute(UserIndex)
        
208             Case eGMCommands.BanChar                 '/BAN
210                 Call HandleBanChar(UserIndex)
        
212             Case eGMCommands.UnbanChar               '/UNBAN
214                 Call HandleUnbanChar(UserIndex)
        
216             Case eGMCommands.NPCFollow               '/SEGUIR
218                 Call HandleNPCFollow(UserIndex)
        
220             Case eGMCommands.SummonChar              '/SUM
222                 Call HandleSummonChar(UserIndex)
        
224             Case eGMCommands.SpawnListRequest        '/CC
226                 Call HandleSpawnListRequest(UserIndex)
        
228             Case eGMCommands.SpawnCreature           'SPA
230                 Call HandleSpawnCreature(UserIndex)
        
232             Case eGMCommands.ResetNPCInventory       '/RESETINV
234                 Call HandleResetNPCInventory(UserIndex)
        
236             Case eGMCommands.CleanWorld              '/LIMPIAR
238                 Call HandleCleanWorld(UserIndex)
        
240             Case eGMCommands.ServerMessage           '/RMSG
242                 Call HandleServerMessage(UserIndex)
        
244             Case eGMCommands.MapMessage              '/MAPMSG
246                 Call HandleMapMessage(UserIndex)
            
248             Case eGMCommands.NickToIP                '/NICK2IP
250                 Call HandleNickToIP(UserIndex)
        
252             Case eGMCommands.IpToNick                '/IP2NICK
254                 Call HandleIPToNick(UserIndex)
        
256             Case eGMCommands.TeleportCreate          '/CT
258                 Call HandleTeleportCreate(UserIndex)
        
260             Case eGMCommands.TeleportDestroy         '/DT
262                 Call HandleTeleportDestroy(UserIndex)
        
268             Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
270                 Call HanldeForceMIDIToMap(UserIndex)
        
272             Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
274                 Call HandleForceWAVEToMap(UserIndex)
        
276             Case eGMCommands.RoyalArmyMessage        '/REALMSG
278                 Call HandleRoyalArmyMessage(UserIndex)
        
280             Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
282                 Call HandleChaosLegionMessage(UserIndex)
        
284             Case eGMCommands.TalkAsNPC               '/TALKAS
286                 Call HandleTalkAsNPC(UserIndex)
        
288             Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
290                 Call HandleDestroyAllItemsInArea(UserIndex)
        
292             Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
294                 Call HandleAcceptRoyalCouncilMember(UserIndex)
        
296             Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
298                 Call HandleAcceptChaosCouncilMember(UserIndex)
        
300             Case eGMCommands.ItemsInTheFloor         '/PISO
302                 Call HandleItemsInTheFloor(UserIndex)
        
304             Case eGMCommands.CouncilKick             '/KICKCONSE
306                 Call HandleCouncilKick(UserIndex)
        
308             Case eGMCommands.SetTrigger              '/TRIGGER
310                 Call HandleSetTrigger(UserIndex)
        
312             Case eGMCommands.AskTrigger              '/TRIGGER with no args
314                 Call HandleAskTrigger(UserIndex)
        
316             Case eGMCommands.BannedIPList            '/BANIPLIST
318                 Call HandleBannedIPList(UserIndex)
        
320             Case eGMCommands.BannedIPReload          '/BANIPRELOAD
322                 Call HandleBannedIPReload(UserIndex)
        
324             Case eGMCommands.BanIP                   '/BANIP
326                 Call HandleBanIP(UserIndex)
        
328             Case eGMCommands.UnbanIP                 '/UNBANIP
330                 Call HandleUnbanIP(UserIndex)
        
332             Case eGMCommands.CreateItem              '/CI
334                 Call HandleCreateItem(UserIndex)
        
336             Case eGMCommands.DestroyItems            '/DEST
338                 Call HandleDestroyItems(UserIndex)
        
340             Case eGMCommands.ChaosLegionKick         '/NOCAOS
342                 Call HandleChaosLegionKick(UserIndex)
        
344             Case eGMCommands.RoyalArmyKick           '/NOREAL
346                 Call HandleRoyalArmyKick(UserIndex)
        
348             Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
350                 Call HandleForceMIDIAll(UserIndex)
        
352             Case eGMCommands.ForceWAVEAll            '/FORCEWAV
354                 Call HandleForceWAVEAll(UserIndex)
        
356             Case eGMCommands.TileBlockedToggle       '/BLOQ
358                 Call HandleTileBlockedToggle(UserIndex)
        
360             Case eGMCommands.KillNPCNoRespawn        '/MATA
362                 Call HandleKillNPCNoRespawn(UserIndex)
        
364             Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
366                 Call HandleKillAllNearbyNPCs(UserIndex)
        
368             Case eGMCommands.LastIP                  '/LASTIP
370                 Call HandleLastIP(UserIndex)
        
372             Case eGMCommands.SystemMessage           '/SMSG
374                 Call HandleSystemMessage(UserIndex)
        
376             Case eGMCommands.CreateNPC               '/ACC
378                 Call HandleCreateNPC(UserIndex)
        
380             Case eGMCommands.CreateNPCWithRespawn    '/RACC
382                 Call HandleCreateNPCWithRespawn(UserIndex)
        
388             Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
390                 Call HandleServerOpenToUsersToggle(UserIndex)
        
392             Case eGMCommands.TurnOffServer           '/APAGAR
394                 Call HandleTurnOffServer(UserIndex)
        
396             Case eGMCommands.TurnCriminal            '/CONDEN
398                 Call HandleTurnCriminal(UserIndex)
        
400             Case eGMCommands.ResetFactions           '/RAJAR
402                 Call HandleResetFactions(UserIndex)
        
404             Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
406                 Call HandleDoBackUp(UserIndex)
        
408             Case eGMCommands.SaveMap                 '/GUARDAMAPA
410                 Call HandleSaveMap(UserIndex)
        
412             Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
414                 Call HandleChangeMapInfoPK(UserIndex)
            
416             Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
418                 Call HandleChangeMapInfoBackup(UserIndex)
        
420             Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
422                 Call HandleChangeMapInfoRestricted(UserIndex)
            
424             Case eGMCommands.ChangeMapInfoLvl
426                 Call HandleChangeMapInfoLvl(UserIndex)
            
428             Case eGMCommands.ChangeMapInfoLimpieza
430                 Call HandleChangeMapInfoLimpieza(UserIndex)
            
432             Case eGMCommands.ChangeMapInfoItems
434                 Call HandleChangeMapInfoItems(UserIndex)

                  Case eGMCommands.ChangeMapExp
                    Call HandleChangeMapInfoExp(UserIndex)
                    
                Case eGMCommands.ChangeMapInfoAttack
                    Call HandleChangeMapInfoAttack(UserIndex)
                    
436             Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
438                 Call HandleChangeMapInfoNoMagic(UserIndex)
        
440             Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
442                 Call HandleChangeMapInfoNoInvi(UserIndex)
        
444             Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
446                 Call HandleChangeMapInfoNoResu(UserIndex)
        
448             Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
450                 Call HandleChangeMapInfoLand(UserIndex)
        
452             Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
454                 Call HandleChangeMapInfoZone(UserIndex)
        
456             Case eGMCommands.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
458                 Call HandleChangeMapInfoStealNpc(UserIndex)
            
460             Case eGMCommands.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
462                 Call HandleChangeMapInfoNoOcultar(UserIndex)
            
464             Case eGMCommands.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
466                 Call HandleChangeMapInfoNoInvocar(UserIndex)
            
468             Case eGMCommands.SaveChars               '/GRABAR
470                 Call HandleSaveChars(UserIndex)
        
476             Case eGMCommands.ChatColor               '/CHATCOLOR
478                 Call HandleChatColor(UserIndex)
        
480             Case eGMCommands.Ignored                 '/IGNORADO
482                 Call HandleIgnored(UserIndex)
            
488             Case eGMCommands.CreatePretorianClan     '/CREARPRETORIANOS
490                 Call HandleCreatePretorianClan(UserIndex)
         
492             Case eGMCommands.RemovePretorianClan     '/ELIMINARPRETORIANOS
494                 Call HandleDeletePretorianClan(UserIndex)
                
496             Case eGMCommands.EnableDenounces         '/DENUNCIAS
498                 Call HandleEnableDenounces(UserIndex)
            
500             Case eGMCommands.ShowDenouncesList       '/SHOW DENUNCIAS
502                 Call HandleShowDenouncesList(UserIndex)
        
504             Case eGMCommands.SetDialog               '/SETDIALOG
506                 Call HandleSetDialog(UserIndex)
            
508             Case eGMCommands.Impersonate             '/IMPERSONAR
510                 Call HandleImpersonate(UserIndex)
            
512             Case eGMCommands.Imitate                 '/MIMETIZAR
514                 Call HandleImitate(UserIndex)
            
516             Case eGMCommands.RecordAdd
518                 Call HandleRecordAdd(UserIndex)
            
520             Case eGMCommands.RecordAddObs
522                 Call HandleRecordAddObs(UserIndex)
            
524             Case eGMCommands.RecordRemove
526                 Call HandleRecordRemove(UserIndex)
            
528             Case eGMCommands.RecordListRequest
530                 Call HandleRecordListRequest(UserIndex)
            
532             Case eGMCommands.RecordDetailsRequest
534                 Call HandleRecordDetailsRequest(UserIndex)
            
536             Case eGMCommands.SearchObj
538                 Call HandleSearchObj(UserIndex)
            
540             Case eGMCommands.SolicitaSeguridad
542                 Call HandleSolicitaSeguridad(UserIndex)
            
544             Case eGMCommands.CheckingGlobal
546                 Call HandleCheckingGlobal(UserIndex)
            
548             Case eGMCommands.CountDown
550                 Call HandleCountDown(UserIndex)
            
552             Case eGMCommands.GiveBackUser
554                 Call HandleGiveBackUser(UserIndex)
            
572             Case eGMCommands.Pro_Seguimiento
574                 Call HandlePro_Seguimiento(UserIndex)
            
576             Case eGMCommands.Events_KickUser
578                 Call HandleEvents_KickUser(UserIndex)
            
580             Case eGMCommands.SendDataUser
582                 Call HandleSendDataUser(UserIndex)
                
584             Case eGMCommands.SearchDataUser
586                 Call HandleSearchDataUser(UserIndex)
                
                  Case eGMCommands.ChangeModoArgentum
                        Call HandleChangeModoArgentum(UserIndex)
                        
                  Case eGMCommands.StreamerBotSetting
                        Call HandleStreamerBotSetting(UserIndex)
                    
                Case eGMCommands.LotteryNew
                    Call HandleLotteryNew(UserIndex)
            End Select

        End With

        Exit Sub

588     Call LogError("Error en GmCommands. Error: " & Err.number & " - " & Err.description & ". Paquete: " & Command)

        '<EhFooter>
        Exit Sub

HandleGMCommands_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGMCommands " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTalk_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 13/01/2010
        '15/07/2009: ZaMa - Now invisible admins talk by console.
        '23/09/2009: ZaMa - Now invisible admins can't send empty chat.
        '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim chat      As String
            Dim ValidChat As Boolean
            Dim PacketCounter As Long
            Dim Packet_ID As Long
       
         
102         ValidChat = True
        
104         chat = Reader.ReadString16()
             
            PacketCounter = Reader.ReadInt32
            Packet_ID = PacketNames.Talk
            
            Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Talk", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
                      
                      
106         ValidChat = Interval_Message(UserIndex)
                
             If Len(chat) >= 300 Then Exit Sub
             
108         If EsGm(UserIndex) Then
110             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eChat, "Dijo: " & chat)
            End If
        
112         Call CheckingOcultation(UserIndex)
        
114         If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then Exit Sub
               End If
        
118         If .flags.Silenciado = 1 Then
120             ValidChat = False
            End If
        
122         If Not PalabraPermitida(LCase$(chat)) Then
124             Call WriteConsoleMsg(UserIndex, "Según la detección automática de insultos, podrías haber insultado. Recuerda que si te sacan una 'FotoDenuncia' irás a la carcel y depende la gravedad podrás recibir baneo completo de cuenta.", FontTypeNames.FONTTYPE_GMMSG)
            End If
        
128         If LenB(chat) <> 0 And ValidChat Then
            
130             If Not (.flags.AdminInvisible = 1) Then
132                 If .flags.Muerto = 1 Then
134                     Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_DEAD_CHAR))
                    Else
136                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(chat, .Char.charindex, 1))
                    End If
                    
                    If Len(chat) >= 3 Then
                        Call WriteAnalyzeText(.Name, chat)
                    End If
                    
                Else

138                 If RTrim(chat) <> "" Then
140                     Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Gm '" & .Name & "'> " & chat, FontTypeNames.FONTTYPE_GM))
                    End If
                End If
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleTalk_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTalk " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleYell_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 13/01/2010 (ZaMa)
        '15/07/2009: ZaMa - Now invisible admins yell by console.
        '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim chat      As String

            Dim UserKey   As Integer

            Dim ValidChat As Boolean

102         ValidChat = True
        
104         chat = Reader.ReadString16()

106         ValidChat = Interval_Message(UserIndex)

108         If EsGm(UserIndex) Then
110             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eChat, "Grito: " & chat)
            End If
            
112         Call CheckingOcultation(UserIndex)

114         If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then Exit Sub
               End If
        
118         If .flags.Silenciado = 1 Then
120             ValidChat = False
            End If
        
122         If ValidChat Then
124             If .flags.Privilegios And PlayerType.User Then
126                 If UserList(UserIndex).flags.Muerto = 1 Then
128                     Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_DEAD_CHAR))
                    Else
130                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(chat, .Char.charindex, 4))
                    End If

                Else

132                 If Not (.flags.AdminInvisible = 1) Then
134                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_GM_YELL))
                    Else
136                     Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Gms> " & chat, FontTypeNames.FONTTYPE_GM))
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleYell_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleYell " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleWhisper_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 03/12/2010
        '28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
        '15/07/2009: ZaMa - Now invisible admins wisper by console.
        '03/12/2010: Enanoh - Agregué susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim chat            As String

            Dim targetUserIndex As Integer

            Dim TargetPriv      As PlayerType

            Dim UserPriv        As PlayerType

            Dim TargetName      As String

            Dim ValidChat       As Boolean

102         ValidChat = True
        
104         TargetName = Reader.ReadString8()
106         chat = Reader.ReadString16()
        
108         ValidChat = Interval_Message(UserIndex)
        
110         UserPriv = .flags.Privilegios
        
        
            If .flags.SlotEvent > 0 Then Exit Sub
                
112         If .flags.Muerto Then
114             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
            Else
                ' Offline?
116             targetUserIndex = NameIndex(TargetName)

118             If targetUserIndex = 0 Then

                    ' Admin?
120                 If EsGmChar(TargetName) Then
122                     Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                        ' Whisperer admin? (Else say nothing)
124                 ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
126                     Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If
                
                    ' Online
                Else
                    ' Privilegios
128                 TargetPriv = UserList(targetUserIndex).flags.Privilegios
                
                    ' Semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
130                 If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (UserPriv And (PlayerType.User Or PlayerType.SemiDios)) <> 0 And Not .flags.EnConsulta Then
                    
                        ' No puede
132                     Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                        ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
134                 ElseIf (UserPriv And PlayerType.User) <> 0 And (Not TargetPriv And PlayerType.User) <> 0 And Not .flags.EnConsulta Then
                    
                        ' No puede
136                     Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                
                        ' En rango? (Los dioses pueden susurrar a distancia)
138                 ElseIf Not EstaPCarea(UserIndex, targetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios)) = 0 Then
                    
                        ' No se puede susurrar a admins fuera de su rango
140                     If (TargetPriv And (PlayerType.User)) = 0 And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
142                         Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    
                            ' Whisperer admin? (Else say nothing)
144                     ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios)) <> 0 Then
146                         Call WriteConsoleMsg(UserIndex, "Estás muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                        End If

                    Else

                        '[GMs]
148                     If UserPriv And (PlayerType.SemiDios) Then
150                         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eChat, "Le susurro a '" & UserList(targetUserIndex).Name & "' " & chat)
                        
                            ' Usuarios a administradores
152                     ElseIf (UserPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
154                         Call Logs_User(UserList(targetUserIndex).Name, eLog.eGm, eLogDescUser.eChat, .Name & " le susurro en consulta: " & chat)
                        End If

156                     If .flags.SlotEvent > 0 Then
158                         If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then ValidChat = False
                        End If
        
160                     If .flags.Silenciado = 1 Then
162                         ValidChat = False
                        End If
                    
164                     If LenB(chat) <> 0 And ValidChat Then
                            ' Dios susurrando a distancia
166                         If Not EstaPCarea(UserIndex, targetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios)) <> 0 Then
                            
168                             Call WriteConsoleMsg(UserIndex, "Susurraste> " & chat, FontTypeNames.FONTTYPE_GM)

170                             Call WriteConsoleMsg(targetUserIndex, "Gm susurra> " & chat, FontTypeNames.FONTTYPE_GM)
                            
172                         ElseIf Not (.flags.AdminInvisible = 1) Then
174                             Call WriteChatPersonalizado(UserIndex, chat, .Char.charindex, 6)
176                             Call WriteChatPersonalizado(targetUserIndex, chat, .Char.charindex, 6)
178                             Call FlushBuffer(targetUserIndex)
                            
                                '[CDT 17-02-2004]
180                             If .flags.Privilegios And (PlayerType.User) Then
182                                 Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & chat, .Char.charindex, vbYellow))
                                End If

                            Else
184                             Call WriteConsoleMsg(UserIndex, "Susurraste> " & chat, FontTypeNames.FONTTYPE_GM)

186                             If UserIndex <> targetUserIndex Then Call WriteConsoleMsg(targetUserIndex, "Gm susurra> " & chat, FontTypeNames.FONTTYPE_GM)
                            
188                             If .flags.Privilegios And (PlayerType.User) Then
190                                 Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(targetUserIndex).Name & "> " & chat, FontTypeNames.FONTTYPE_GM))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleWhisper_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleWhisper " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '11/19/09 Pato - Now the class bandit can walk hidden.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    Dim dummy       As Long

    Dim TempTick    As Long

    Dim Heading     As eHeading
    
    Dim MaxTimeWalk As Integer
        
    Dim PacketCount As Long
        
    With UserList(UserIndex)
        
        Heading = Reader.ReadInt()
        PacketCount = Reader.ReadInt32
            
        Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Walk), .PacketTimers(PacketNames.Walk), .MacroIterations(PacketNames.Walk), UserIndex, "Walk", PacketTimerThreshold(PacketNames.Walk), MacroIterations(PacketNames.Walk))
            
        If .flags.Muerto Then
            MaxTimeWalk = 36
        Else
            MaxTimeWalk = 30

        End If
        
        If .flags.Paralizado = 0 Then
            
            If .flags.Meditando Then
                
                ' Probabilidad de subir un % de maná al moverse
                If RandomNumber(1, 100) <= 20 Then

                    Dim Mana As Long

                    Mana = Porcentaje(.Stats.MaxMan, Porcentaje(Balance.PorcentajeRecuperoMana, 50 + .Stats.UserSkills(eSkill.Magia) * 0.5))

                    If Mana <= 0 Then Mana = 1
                    
                    If .Stats.MinMan + Mana >= .Stats.MaxMan Then
                        .Stats.MinMan = .Stats.MaxMan
                    Else
                        .Stats.MinMan = .Stats.MinMan + Mana
                    End If
                    
                    Call WriteUpdateMana(UserIndex)

                End If
                
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                UserList(UserIndex).Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))

            End If
            
            Dim CurrentTick As Long

            CurrentTick = GetTime
        
            'Prevent SpeedHack (refactored by WyroX)
            If .Char.speeding > 0 Then

                Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single

                ElapsedTimeStep = CurrentTick - .Counters.LastStep
                MinTimeStep = IntervaloCaminar / .Char.speeding
                DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep

                If DeltaStep > 0 Then
                
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                
                    If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                        'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administración Â» Posible uso de SpeedHack del usuario " & .name & ".", e_FontTypeNames.FONTTYPE_SERVER))
                        Call WritePosUpdate(UserIndex)
                        Exit Sub

                    End If

                Else
                
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5

                    If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

                End If

            End If
            
            ' @ En la daga rusa no te podes mover [El chequeo está en el cliente, pero al empezar se mueven como retrasados]
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = eModalityEvent.DagaRusa Then
                    If Events(.flags.SlotEvent).Run Then Exit Sub

                End If

            End If
            
            'Move user
            If MoveUserChar(UserIndex, Heading) Then
                ' Save current step for anti-sh
                .Counters.LastStep = CurrentTick
                         
                If UserIndex <> StreamerBot.Active And StreamerBot.Active > 0 Then
                   ' If StrComp(StreamerBot.LastTarget, UCase$(.Name)) = 0 Then
                       
                    'End If
                    
                   ' If StrComp(StreamerBot.LastTarget, UCase$(.Name)) = 0 Then
                        ' If MoveUserChar(StreamerBot.Active, Heading) Then
                         '   Call WriteForceCharMove(StreamerBot.Active, Heading)
                        'End If
                    'End If

                End If

                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)

                End If

                'If exiting, cancel
                Call CancelExit(UserIndex)
                
                'Esta usando el /HOGAR, no se puede mover
                If .flags.Traveling = 1 Then
                    .flags.Traveling = 0
                    .Counters.goHome = 0
                    Call WriteConsoleMsg(UserIndex, "Has cancelado el viaje a casa.", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Else
                .Counters.LastStep = 0
                Call WritePosUpdate(UserIndex)

            End If
            
        Else    'paralized

            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                            
            Dim HunterInPosValid As Boolean
                
            ' @ Cazadores y sus capuchas weonas
            '  If .Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
            If .Invent.CascoEqpObjIndex > 0 Then
                If ObjData(.Invent.CascoEqpObjIndex).Oculto > 0 Then

                    ' Si está en el rango permitido desde que se ocultó, puede moverse libre.
                    ' Esta dentro del rango permitido
                    If Distance(.Pos.X, .Pos.Y, .PosOculto.X, .PosOculto.Y) <= ObjData(.Invent.CascoEqpObjIndex).Oculto Then
                        HunterInPosValid = True

                    End If
                
                End If

            End If

            ' End If
                
            If .Clase <> eClass.Thief And Not HunterInPosValid Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
            
                If .flags.Navegando = 0 Then

                    'If not under a spell effect, show char
                    If .flags.Invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.charindex, False)

                    End If

                End If

            End If

        End If
        
        .Counters.PiqueteC = 0
        Call Guilds_UpdatePosition(UserIndex)

    End With

End Sub

Public Function Check_UserBlocked(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    On Error GoTo ErrHandler
    
    If MapData(Map, X - 1, Y).Blocked = 0 And _
        MapData(Map, X - 1, Y).NpcIndex = 0 And _
        MapData(Map, X - 1, Y).UserIndex = 0 And _
        MapData(Map, X - 1, Y).TileExit.Map = 0 Then
            
        Check_UserBlocked = False
        Exit Function
    End If
    
    If MapData(Map, X + 1, Y).Blocked = 0 And _
        MapData(Map, X + 1, Y).NpcIndex = 0 And _
        MapData(Map, X + 1, Y).UserIndex = 0 And _
        MapData(Map, X + 1, Y).TileExit.Map = 0 Then
            
        Check_UserBlocked = False
        Exit Function
    End If
    
    If MapData(Map, X, Y - 1).Blocked = 0 And _
        MapData(Map, X, Y - 1).NpcIndex = 0 And _
        MapData(Map, X, Y - 1).UserIndex = 0 And _
        MapData(Map, X, Y - 1).TileExit.Map = 0 Then
        Check_UserBlocked = False
        Exit Function
    End If
    
    If MapData(Map, X, Y + 1).Blocked = 0 And _
        MapData(Map, X, Y + 1).NpcIndex = 0 And _
        MapData(Map, X, Y + 1).UserIndex = 0 And _
        MapData(Map, X, Y + 1).TileExit.Map = 0 Then
            
        Check_UserBlocked = False
        Exit Function
    End If
    
    Check_UserBlocked = True
    Exit Function
ErrHandler:
    
End Function

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleRequestPositionUpdate_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 29/10/2021
        '
        '***************************************************
    
        Dim Pos    As WorldPos

        Dim OldPos As WorldPos
    
100     With UserList(UserIndex)
102         Pos = .Pos
        
104         If .flags.SlotReto = 0 And .flags.SlotEvent = 0 And .flags.SlotFast = 0 And .flags.Desafiando = 0 And Not MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = ZONAPELEA Then
            
106             If Check_UserBlocked(Pos.Map, Pos.X, Pos.Y) Then
108                 Call ClosestStablePos(Pos, Pos)

                    If Pos.X <> 0 And .Pos.Y <> 0 Then
110                     Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, True)

                    End If

                End If

            End If
        
        End With
    
112     Call WritePosUpdate(UserIndex)
        '<EhFooter>
        Exit Sub

HandleRequestPositionUpdate_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestPositionUpdate " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleAttack_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 13/01/2010
        'Last Modified By: ZaMa
        '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
        '13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
        '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
        '***************************************************
    
100     With UserList(UserIndex)
          
            Dim PacketCounter As Long

            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long

            Packet_ID = PacketNames.Attack
            
            Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Attack", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            
102         If .flags.GmSeguidor > 0 Then

                Dim Temp As Long, TiempoActual As Long

104             TiempoActual = GetTime
106             Temp = TiempoActual - .interval(0).IAttack
                    
108             Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 4, Temp, .flags.MenuCliente)
                    
110             .interval(0).IAttack = TiempoActual

            End If
        
            'If dead, can't attack
112         If .flags.Muerto = 1 Then
114             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'If user meditates, can't attack
116         If .flags.Meditando Then

                Exit Sub

            End If
        
            'If equiped weapon is ranged, can't attack this way
118         If .Invent.WeaponEqpObjIndex > 0 Then
120             If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
122                 Call WriteConsoleMsg(UserIndex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            End If
        
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).flags.Meditando = False
                UserList(UserIndex).Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))

            End If
            
            'If exiting, cancel
124         Call CancelExit(UserIndex)
        
            'Attack!
126         Call UsuarioAtaca(UserIndex)
        
            'Now you can be atacked
128         .flags.NoPuedeSerAtacado = False
        
            'I see you...
130         If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
132             .flags.Oculto = 0
134             .Counters.TiempoOculto = 0
            
136             If .flags.Navegando = 0 Then

138                 If .flags.Invisible = 0 Then
140                     Call UsUaRiOs.SetInvisible(UserIndex, .Char.charindex, False)
142                     Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If
     
        End With

        '<EhFooter>
        Exit Sub

HandleAttack_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAttack " & "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePickUp_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 07/25/09
        '02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
        '***************************************************
    
100     With UserList(UserIndex)
        
            'If dead, it can't pick up objects
102         If .flags.Muerto = 1 Then Exit Sub
        
            'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
104         If .flags.Comerciando Then Exit Sub
        
110         Call GetObj(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandlePickUp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePickUp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSafeToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
            
            If .Faction.Status = r_Armada Then
                Call WriteConsoleMsg(UserIndex, "Tu facción no te permite quitar el seguro. Por favor dirigete al Rey de Banderbill y abandona la facción.", FONTTYPE_WARNING)
                Exit Sub
            End If
            
102         If .flags.Seguro Then
104             Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
            Else
106             Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
            End If
        
108         .flags.Seguro = Not .flags.Seguro
        End With

        '<EhFooter>
        Exit Sub

HandleSafeToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSafeToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleResuscitationToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Rapsodius
        'Creation Date: 10/10/07
        '***************************************************
100     With UserList(UserIndex)
        
102         .flags.SeguroResu = Not .flags.SeguroResu
        
104         If .flags.SeguroResu Then
106             Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
            Else
108             Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleResuscitationToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleResuscitationToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleDragToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDragToggle_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
102         .flags.DragBlocked = Not .flags.DragBlocked
        
104         If .flags.DragBlocked Then
106             Call WriteMultiMessage(UserIndex, eMessages.DragSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
            Else
108             Call WriteMultiMessage(UserIndex, eMessages.DragSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
            End If
    
        End With

        '<EhFooter>
        Exit Sub

HandleDragToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDragToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestAtributes_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     Call WriteAttributes(UserIndex)
        '<EhFooter>
        Exit Sub

HandleRequestAtributes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestAtributes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestSkills_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     Call WriteSendSkills(UserIndex)
        '<EhFooter>
        Exit Sub

HandleRequestSkills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestSkills " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestMiniStats_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     Call WriteMiniStats(UserIndex, UserIndex)
        '<EhFooter>
        Exit Sub

HandleRequestMiniStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestMiniStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCommerceEnd_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        'User quits commerce mode
100     UserList(UserIndex).flags.Comerciando = False
102     Call WriteCommerceEnd(UserIndex)
        '<EhFooter>
        Exit Sub

HandleCommerceEnd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCommerceEnd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUserCommerceEnd_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 11/03/2010
        '11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
        '***************************************************
100     With UserList(UserIndex)
        
            'Quits commerce mode with user
102         If .ComUsu.DestUsu > 0 Then
104             If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
106                 Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
108                 Call FinComerciarUsu(.ComUsu.DestUsu)
                
                    'Send data in the outgoing buffer of the other user
110                 Call FlushBuffer(.ComUsu.DestUsu)
                End If
            End If
        
112         Call FinComerciarUsu(UserIndex)
114         Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
        End With

        '<EhFooter>
        Exit Sub

HandleUserCommerceEnd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUserCommerceEnd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUserCommerceConfirm_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        '
        '***************************************************

        'Validate the commerce
100     If PuedeSeguirComerciando(UserIndex) Then
            'Tell the other user the confirmation of the offer
102         Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
104         UserList(UserIndex).ComUsu.Confirmo = True
        End If
    
        '<EhFooter>
        Exit Sub

HandleUserCommerceConfirm_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUserCommerceConfirm " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCommerceChat_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 03/12/2009
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim chat As String
        
102         chat = Reader.ReadString8()
        
104         If LenB(chat) <> 0 Then
106             If PuedeSeguirComerciando(UserIndex) Then
                
108                 chat = UserList(UserIndex).Name & "> " & chat
110                 Call WriteCommerceChat(UserIndex, chat, FontTypeNames.FONTTYPE_PARTY)
112                 Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, chat, FontTypeNames.FONTTYPE_PARTY)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleCommerceChat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCommerceChat " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBankEnd_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
            'User exits banking mode
102         .flags.Comerciando = False
104         Call WriteBankEnd(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleBankEnd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBankEnd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUserCommerceOk_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        'Trade accepted
100     Call AceptarComercioUsu(UserIndex)
        '<EhFooter>
        Exit Sub

HandleUserCommerceOk_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUserCommerceOk " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUserCommerceReject_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim otherUser As Integer
    
100     With UserList(UserIndex)
        
102         otherUser = .ComUsu.DestUsu
        
            'Offer rejected
104         If otherUser > 0 Then
106             If UserList(otherUser).flags.UserLogged Then
108                 Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
110                 Call FinComerciarUsu(otherUser)
                
                    'Send data in the outgoing buffer of the other user
112                 Call FlushBuffer(otherUser)
                End If
            End If
        
114         Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
116         Call FinComerciarUsu(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleUserCommerceReject_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUserCommerceReject " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDrop_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 07/25/09
        '07/25/09: Marco - Agregué un checkeo para patear a los usuarios que tiran items mientras comercian.
        '***************************************************
    
        Dim Slot    As Byte

        Dim Amount  As Integer
    
100     With UserList(UserIndex)

102         Slot = Reader.ReadInt()
104         Amount = Reader.ReadInt()

106         If Not Interval_Drop(UserIndex) Then Exit Sub
        
            'low rank admins can't drop item. Neither can the dead nor those sailing.
108         If .flags.Navegando = 1 Or .flags.Muerto = 1 Or .flags.Montando = 1 Or .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Sub

            'If the user is trading, he can't drop items => He's cheating, we kick him.
116         If .flags.Comerciando Then Exit Sub
        
118         If Slot = FLAGORO + 1 Then

                Exit Sub

                'If Amount > 10000 Then Exit Sub 'Don't drop too much gold
                'If (.Stats.Eldhir - Amount) < 0 Then Exit Sub
            
                'Dim Pos As WorldPos
                'Dim Obj As Obj
            
                'Obj.ObjIndex = 1246
                'Obj.Amount = Amount
            
                'TirarItemAlPiso .Pos, Obj
                ' .Stats.Eldhir = .Stats.Eldhir - Amount
                'Call WriteUpdateDsp(UserIndex)
            
                'Are we dropping gold or other items??
120         ElseIf Slot = FLAGORO Then
                If Amount > 10000 Then Exit Sub 'Don't drop too much gold
                If (.Stats.Gld - Amount) < 0 Then Exit Sub
            
                Dim Pos As WorldPos
                Dim Obj As Obj
            
                Obj.ObjIndex = iORO
                Obj.Amount = Amount
            
                TirarItemAlPiso .Pos, Obj
                 .Stats.Gld = .Stats.Gld - Amount
                Call WriteUpdateGold(UserIndex)
                Exit Sub
            Else

                'Only drop valid slots
122             If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
124                 If .Invent.Object(Slot).ObjIndex = 0 Then

                        Exit Sub

                    End If
                
126                 Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleDrop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDrop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCastSpell_Err
        '</EhHeader>


        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
        '***************************************************
100     With UserList(UserIndex)
        
            Dim Spell As Byte
        
102         Spell = Reader.ReadInt()
104         Reader.ReadInt16
106         Reader.ReadInt8
        
108         If Not IntervaloPermiteCastear(UserIndex, True) Then Exit Sub    'Nuevo intervalo de casteo.
        
110         If .flags.Muerto = 1 Then
112             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
114         If .flags.MenuCliente <> 255 And .flags.MenuCliente <> 1 Then

                Exit Sub

            End If
        
            'Now you can be atacked
116         .flags.NoPuedeSerAtacado = False
        
118         If Spell < 1 Then
120             .flags.Hechizo = 0

                Exit Sub

122         ElseIf Spell > MAXUSERHECHIZOS Then
124             .flags.Hechizo = 0

                Exit Sub

            End If
        
126         .flags.Hechizo = .Stats.UserHechizos(Spell)

128         If Hechizos(.flags.Hechizo).AutoLanzar = 1 Then

                'Check bow's interval
130             If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                'Check attack-spell interval
132             If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
                'Check Magic interval
134             If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub

136             .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
138             Call LanzarHechizo(.flags.Hechizo, UserIndex)
140             .flags.Hechizo = 0

            End If
            
        End With

        '<EhFooter>
        Exit Sub

HandleCastSpell_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCastSpell " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLeftClick_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        Dim X       As Byte

        Dim Y       As Byte
        
100     X = Reader.ReadInt()
102     Y = Reader.ReadInt()
        
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.LeftClick
            
            
            With UserList(UserIndex)
            Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "LeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            
104     Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
            
            End With
           
 
        '<EhFooter>
        Exit Sub

HandleLeftClick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLeftClick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDoubleClick_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        Dim X       As Byte

        Dim Y       As Byte
        
        Dim Tipo As Byte
        
100     X = Reader.ReadInt8()
102     Y = Reader.ReadInt8()
          Tipo = Reader.ReadInt8
          
104     Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y, Tipo)

        '<EhFooter>
        Exit Sub

HandleDoubleClick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDoubleClick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RightClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRightClick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRightClick_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 10/05/2011
        '
        '***************************************************
        
        Dim X       As Byte

        Dim Y       As Byte
        
        Dim MouseX As Long
        Dim MouseY As Long

        Dim UserKey As Integer
        
100     X = Reader.ReadInt8()
102     Y = Reader.ReadInt8()

        
            MouseX = Reader.ReadInt32()
            MouseY = Reader.ReadInt32()
            
            
104     Call Extra.ShowMenu(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    
        '<EhFooter>
        Exit Sub

HandleRightClick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRightClick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleWork_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 13/01/2010 (ZaMa)
        '13/01/2010: ZaMa - El pirata se puede ocultar en barca
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Skill   As eSkill

            Dim UserKey As Integer
        
102         Skill = Reader.ReadInt()
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.Work
            
            
106         If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
            'If exiting, cancel
108         Call CancelExit(UserIndex)
        
110         Select Case Skill
        
                Case Robar, Magia, Domar
112                 Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill)
                
114             Case Ocultarse
                    Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Ocultar", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
                    
                    ' Verifico si se peude ocultar en este mapa
116                 If MapInfo(.Pos.Map).OcultarSinEfecto = 1 Then
118                     Call WriteConsoleMsg(UserIndex, "¡Ocultarse no funciona aquí!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    If .flags.SlotFast > 0 Then
                        If RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                            Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el ocultamiento.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                            Exit Sub
    
                        End If
                    End If
                
120                 If .flags.SlotEvent > 0 Then
122                     If Events(.flags.SlotEvent).config(eConfigEvent.eOcultar) = 0 Then
124                         Call WriteConsoleMsg(UserIndex, "Ocultar no está permitido aquí! Retirate de la Zona del Evento si deseas esconderte entre las sombras.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                
126                 If .flags.EnConsulta Then
128                     Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
130                 'If .Stats.MaxMan > 0 Then
132                    ' Call WriteConsoleMsg(UserIndex, "No tienes el conocimiento para ocultarte entre las sombras.", FontTypeNames.FONTTYPE_INFO)
                
                     '   Exit Sub
                  '  End If
                
134                 If Power.UserIndex = UserIndex Then
136                     Call WriteConsoleMsg(UserIndex, "¿A que seguro eres un cazador ah!? ¡Plantate!", FontTypeNames.FONTTYPE_INFO)
                
                        Exit Sub
                    End If
                
138                 If .flags.Navegando = 1 Or .flags.Montando = 1 Or .flags.Mimetizado = 1 Or .flags.Invisible = 1 Then
                        '[CDT 17-02-2004]
140                     If Not .flags.UltimoMensaje = 3 Then
142                         Call WriteConsoleMsg(UserIndex, "No puedes ocultarte en este momento .", FontTypeNames.FONTTYPE_INFO)
144                         .flags.UltimoMensaje = 3
                        End If

                        '[/CDT]
                        Exit Sub
                    End If
                
                
154                 If .flags.Oculto = 1 Then

                        '[CDT 17-02-2004]
156                     If Not .flags.UltimoMensaje = 2 Then
158                         Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
160                         .flags.UltimoMensaje = 2
                        End If

                        '[/CDT]
                        Exit Sub

                    End If
                
162                 Call DoOcultarse(UserIndex)
                
            End Select
        
        End With

        '<EhFooter>
        Exit Sub

HandleWork_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleWork " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleInitCrafting_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/01/2010
        '
        '***************************************************
    
        Dim TotalItems    As Long

        Dim ItemsPorCiclo As Integer
    
100     With UserList(UserIndex)
        
102         TotalItems = Reader.ReadInt
104         ItemsPorCiclo = Reader.ReadInt
        
106         If TotalItems > 0 Then
            
108             .Construir.cantidad = TotalItems
110             .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleInitCrafting_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleInitCrafting " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleUseItem_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Slot           As Byte

            Dim SecondaryClick As Byte

            Dim Value          As Long

            Dim UserKey        As Integer
        
            Dim Key            As Integer

            Dim PacketCounter  As Long

            Dim Packet_ID      As Long

102         Slot = Reader.ReadInt8()
104         SecondaryClick = Reader.ReadInt8()
106         Value = Reader.ReadInt32()
     
            PacketCounter = Reader.ReadInt32
            
114         If Slot <= .CurrentInventorySlots And Slot > 0 Then
116             If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
            Else
118             Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " hizo algo raro al usar objetos")
120             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " hizo algo raro al usar objetos", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                Exit Sub

            End If
            
122         If .flags.Meditando Then Exit Sub
124         If .flags.Comerciando Then Exit Sub
        
126         If SecondaryClick And .flags.MenuCliente = 1 Then Exit Sub
128         If .flags.LastSlotClient <> 255 And Slot <> .flags.LastSlotClient Then Exit Sub
              
            If SecondaryClick Then
                
 
                If (GetTime - .TimeUseClicInitial) >= 1000 Then
                    If .TimeUseClic >= 5 Then
                        Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " está utilizando más de 4 doble-clics")
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " está utilizando más de 4 doble-clics", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                    End If
                        
                    .TimeUseClic = 0
                    .TimeUseClicInitial = GetTime
                    Exit Sub
                Else
                    .TimeUseClic = .TimeUseClic + 1
                End If
                
                
                Packet_ID = PacketNames.UseItem
                Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItem", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            Else
                Packet_ID = PacketNames.UseItemU
                Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItemU", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            End If
            
130         Call UseInvItem(UserIndex, Slot, SecondaryClick, Value)

        End With
            
        '<EhFooter>
        Exit Sub

HandleUseItem_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUseItem " & "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItemTwo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUseItemTwo_Err
        '</EhHeader>

      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
    
100   With UserList(UserIndex)
        
          Dim Slot           As Byte

          Dim SecondaryClick As Byte

          Dim Value          As Long

          Dim UserKey        As Integer
        
          Dim Key            As Integer
               

102       Slot = Reader.ReadInt8()
104       SecondaryClick = Reader.ReadInt8()
106       Value = Reader.ReadInt32()

114       If Slot <= .CurrentInventorySlots And Slot > 0 Then
116           If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
          Else
118           Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " hizo algo raro al usar objetos")
120           Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " hizo algo raro al usar objetos", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
              Exit Sub
          End If
        
122       If .flags.Meditando Then Exit Sub
124       If .flags.Comerciando Then Exit Sub
        
126       If SecondaryClick And .flags.MenuCliente = 1 Then Exit Sub
128       If .flags.LastSlotClient <> 255 And Slot <> .flags.LastSlotClient Then Exit Sub
            
            If PacketUseItem <> ClientPacketID.UseItemTwo Then
                Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " utilizo un paquete guardado")
           Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " utilizo un paquete guardado", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
            End If
            
130       Call UseInvItem(UserIndex, Slot, SecondaryClick, Value)

      End With
            
        '<EhFooter>
        Exit Sub

HandleUseItemTwo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUseItemTwo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCraftBlacksmith_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        Dim QuestIndex As Integer
    
100     QuestIndex = Reader.ReadInt16()
    
102     If QuestIndex < 1 Or QuestIndex > NumQuests Then Exit Sub
104     If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        

        '<EhFooter>
        Exit Sub

HandleCraftBlacksmith_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCraftBlacksmith " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleWorkLeftClick_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 14/01/2010 (ZaMa)
        '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
        '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
        '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim X           As Byte

            Dim Y           As Byte

            Dim Skill       As eSkill

            Dim DummyInt    As Integer

            Dim tU          As Integer   'Target user

            Dim tN          As Integer   'Target NPC
        
            Dim WeaponIndex As Integer
        
            Dim Key         As Integer
            
            Dim MouseX      As Long
            
            Dim MouseY      As Long
            
102         X = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
        
106         Skill = Reader.ReadInt8()
108         MouseX = Reader.ReadInt8
110         MouseY = Reader.ReadInt16
              
            Dim PacketCounter As Long

            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long

            Packet_ID = PacketNames.WorkLeftClick
            
112         If (.flags.Muerto = 1 And Skill <> TeleportInvoker) Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
                
            '  If .Clase <> eClass.Worker Then
            ' UpdatePointer UserIndex, .flags.MenuCliente, X, Y, "Click to Win"

            '  End If
              
114         If Not InRangoVision(UserIndex, X, Y) Then
116             Call WritePosUpdate(UserIndex)

                Exit Sub

            End If
            
            If .flags.Meditando Then
                .flags.Meditando = False
                .Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))

            End If
            
            Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "WorkLeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            
            'If exiting, cancel
118         Call CancelExit(UserIndex)
        
120         Select Case Skill

                Case eSkill.Proyectiles
                
                    'Check attack interval
122                 If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

                    'Check Magic interval
124                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub

                    'Check bow's interval
126                 If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
128                 Call LanzarProyectil(UserIndex, X, Y)
                            
130             Case eSkill.Magia

                    'Check the map allows spells to be casted.
132                 If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
134                     Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía.", FontTypeNames.FONTTYPE_FIGHT)

                        Exit Sub

                    End If
                
                    'Target whatever is in that tile
136                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                    'If it's outside range log it and exit
138                 If Abs(.Pos.X - X) > RANGO_VISION_x Or Abs(.Pos.Y - Y) > RANGO_VISION_y Then
140                     Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IpAddress & " a la posición (" & .Pos.Map & "/" & X & "/" & Y & ")")

                        Exit Sub

                    End If
                
                    'Check bow's interval
142                 If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                    'Check attack-spell interval
144                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
                    'Check Magic interval
146                 If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
                 
                    'Check intervals and cast
148                 If .flags.Hechizo > 0 Then
                          If Hechizos(.flags.Hechizo).AutoLanzar = 1 Then Exit Sub ' Anti hack
150                     .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
152                     Call LanzarHechizo(.flags.Hechizo, UserIndex)
154                     .flags.Hechizo = 0
                    Else
156                     Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

                    End If
            
158             Case eSkill.Robar

160                 If .Clase <> eClass.Thief Then
162                     Call WriteConsoleMsg(UserIndex, "¡Tu no puedes robar!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                        
                    'Does the map allow us to steal here?
164                 If MapInfo(.Pos.Map).Pk Then
                    
                        'Check interval
166                     If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                        'Target whatever is in that tile
168                     Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
170                     tU = .flags.TargetUser
                    
172                     If tU > 0 And tU <> UserIndex Then

                            'Can't steal administrative players
174                         If UserList(tU).flags.Privilegios And PlayerType.User Then
176                             If UserList(tU).flags.Muerto = 0 Then
178                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 4 Then
180                                     Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                                        Exit Sub

                                    End If
                                 
                                    '17/09/02
                                    'Check the trigger
182                                 If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
184                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)

                                        Exit Sub

                                    End If
                                 
186                                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
188                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)

                                        Exit Sub

                                    End If
                                 
190                                 Call DoRobar(UserIndex, tU)

                                End If

                            End If

                        Else
192                         Call WriteConsoleMsg(UserIndex, "¡No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
194                     Call WriteConsoleMsg(UserIndex, "¡No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)

                    End If

196             Case eSkill.Domar
                    'Modificado 25/11/02
                    'Optimizado y solucionado el bug de la doma de
                    'criaturas hostiles.
                
                    'Target whatever is that tile
198                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
200                 tN = .flags.TargetNPC
                
202                 If tN > 0 Then
204                     If Npclist(tN).flags.Domable > 0 Then
206                         If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
208                             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                                Exit Sub

                            End If
                        
210                         If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
212                             Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)

                                Exit Sub

                            End If
                        
                            'mMascotas.Mascotas_AddNew UserIndex, tN
                            'Call DoDomar(UserIndex, tN)
                        Else
214                         Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
216                     Call WriteConsoleMsg(UserIndex, "¡No hay ninguna criatura allí!", FontTypeNames.FONTTYPE_INFO)

                    End If
           
218             Case eSkill.Pesca
                
224                 WeaponIndex = .Invent.WeaponEqpObjIndex

226                 If WeaponIndex = 0 Then Exit Sub
                
                    'Check interval
228                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                    'Basado en la idea de Barrin
                    'Comentario por Barrin: jah, "basado", caradura ! ^^
230                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
232                     Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
234                 If HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
236                     If Abs(.Pos.X - .flags.TargetX) + Abs(.Pos.Y - .flags.TargetY) > 6 Then
238                         Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para sacar peces.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                                 
240                     Select Case WeaponIndex

                            Case CAÑA_PESCA, RED_PESCA
242                             Call DoPescar(UserIndex, WeaponIndex)
 
244                         Case Else

                                Exit Sub    'Invalid item!

                        End Select
                    
                        'Play sound!
246                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PESCAR, .Pos.X, .Pos.Y, .Char.charindex))
                    Else
248                     Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)

                    End If

250             Case eSkill.Mineria

252                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
254                 WeaponIndex = .Invent.WeaponEqpObjIndex
                                
256                 If WeaponIndex = 0 Then Exit Sub
                
258                 If (WeaponIndex <> PIQUETE_MINERO) Then

                        ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                        Exit Sub

                    End If
                
                    'Target whatever is in the tile
260                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
262                 DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
264                 If DummyInt > 0 Then

                        'Check distance
266                     If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
268                         Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                    
                        '¿Hay un yacimiento donde clickeo?
270                     If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
272                         Call DoMineria(UserIndex)
                        Else
274                         Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
276                     Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)

                    End If
                  
278             Case TeleportInvoker 'UGLY!!! This is a constant, not a skill!!

                    'Check interval
280                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                          
                    'Validate other items
                    If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then

                        Exit Sub

                    End If
                    
                    If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <= 0 Then Exit Sub ' @@ No se si puede pasar
                    
                    Call Teleports_AddNew(UserIndex, .Invent.Object(.flags.TargetObjInvSlot).ObjIndex, .Pos.Map, X, Y)

282             Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                
                    'Check interval
284                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                    'Check there is a proper item there
286                 If .flags.TargetObj > 0 Then
288                     If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then

                            'Validate other items
290                         If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then

                                Exit Sub

                            End If
                        
                            ''chequeamos que no se zarpe duplicando oro
292                         If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
294                             If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
296                                 Call WriteConsoleMsg(UserIndex, "No tienes más minerales.", FontTypeNames.FONTTYPE_INFO)

                                    Exit Sub

                                End If
                            
                                ''FUISTE
298                             Call Protocol.Kick(UserIndex)

                                Exit Sub

                            End If

300                         If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
302                             Call FundirMineral(UserIndex)
304                         ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then

                                ' Call FundirArmas(UserIndex)
                            End If

                        Else
306                         Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
308                     Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)

                    End If

310             Case eSkill.Talar

                    'Check interval
316                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
318                 WeaponIndex = .Invent.WeaponEqpObjIndex
                
320                 If WeaponIndex = 0 Then
                    
322                     Call WriteConsoleMsg(UserIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
324                 If WeaponIndex <> HACHA_LEÑADOR Then

                        ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                        Exit Sub

                    End If
                
326                 DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
328                 If DummyInt > 0 Then
330                     If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
332                         Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                    
                        'Barrin 29/9/03
334                     If .Pos.X = X And .Pos.Y = Y Then
336                         Call WriteConsoleMsg(UserIndex, "No puedes talar desde allí.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                    
                        '¿Hay un arbol donde clickeo?
338                     If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
340                         If WeaponIndex = HACHA_LEÑADOR Then
                            
                                Dim Objeto As Integer

342                             Objeto = ObjData(DummyInt).ArbolItem
                            
344                             If Objeto = 0 Then
346                                 Call WriteConsoleMsg(UserIndex, "El árbol no posee leños suficientes para poder arrojar.", FontTypeNames.FONTTYPE_INFO)
                                Else
348                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_TALAR, .Pos.X, .Pos.Y, .Char.charindex))
350                                 Call DoTalar(UserIndex, Objeto)

                                End If
                            
                            Else
352                             Call WriteConsoleMsg(UserIndex, "No has podido extraer leña. Comprueba los conocimientos necesarios.", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    Else
354                     Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)

                    End If
            
            End Select

        End With

        '<EhFooter>
        Exit Sub

HandleWorkLeftClick_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWorkLeftClick " & "at line " & Erl

        '</EhFooter>
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSpellInfo_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim spellSlot As Byte

            Dim Spell     As Integer
        
102         spellSlot = Reader.ReadInt()
        
            'Validate slot
104         If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
106             Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Validate spell in the slot
108         Spell = .Stats.UserHechizos(spellSlot)

110         If Spell > 0 And Spell < NumeroHechizos + 1 Then

112             With Hechizos(Spell)
                    'Send information
114                 Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf & "Nombre:" & .Nombre & vbCrLf & "Descripción:" & .Desc & vbCrLf & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf & "Maná necesario: " & .ManaRequerido & vbCrLf & "Energía necesaria: " & .StaRequerido & vbCrLf & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
                End With

            End If

        End With

        '<EhFooter>
        Exit Sub

HandleSpellInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSpellInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEquipItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim itemSlot As Byte
        
102         itemSlot = Reader.ReadInt()
        
            'Dead users can't equip items
104         If .flags.Muerto = 1 Then Exit Sub
        
            'Validate item slot
106         If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
108         If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
                If Not Interval_Equipped(UserIndex) Then Exit Sub
110         Call EquiparInvItem(UserIndex, itemSlot)
        End With

        '<EhFooter>
        Exit Sub

HandleEquipItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEquipItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 06/28/2008
    'Last Modified By: NicoNZ
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    ' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
    '***************************************************

    With UserList(UserIndex)
        
        Dim Heading As eHeading

        Dim posX    As Integer

        Dim posY    As Integer
        Dim PacketCounter As Long
           Heading = Reader.ReadInt()
        PacketCounter = Reader.ReadInt32
                        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.ChangeHeading
            
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "ChangeHeading", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        
     
        
        ' Las clases con maná no se pueden mover.
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 1 Then
            If .Stats.MaxMan <> 0 Then Exit Sub
            
        Else

            If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then

                Exit Sub

            End If

        End If
               
        'If .flags.Paralizado = 1 And .flags.Inmovilizado = 1 Then

        'Select Case Heading

        'Case eHeading.NORTH
        '  posY = -1

        ' Case eHeading.EAST
        '   posX = 1

        ' Case eHeading.SOUTH
        '   posY = 1

        ' Case eHeading.WEST
        '    posX = -1
        ' End Select
            
        ' If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then

        '   Exit Sub

        ' End If
        ' End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeHeading(.Char.charindex, .Char.Heading))

        End If

    End With

End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleModifySkills_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 11/19/09
        '11/19/09: Pato - Adapting to new skills system.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim i                      As Long

            Dim Count                  As Integer

            Dim Points(1 To NUMSKILLS) As Byte
        
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
102         For i = 1 To NUMSKILLS
104             Points(i) = Reader.ReadInt()
            
106             If Points(i) < 0 Then
108                 Call LogHackAttemp(.Name & " IP:" & .IpAddress & " trató de hackear los skills.")
110                 .Stats.SkillPts = 0
112                 Call Protocol.Kick(UserIndex)

                    Exit Sub

                End If
            
114             Count = Count + Points(i)
116         Next i
        
118         If Count > .Stats.SkillPts Then
120             Call LogHackAttemp(.Name & " IP:" & .IpAddress & " trató de hackear los skills.")
122             Call Protocol.Kick(UserIndex)
                Exit Sub

            End If
        
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
124         .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        
126         With .Stats

128             For i = 1 To NUMSKILLS

130                 If Points(i) > 0 Then
132                     .SkillPts = .SkillPts - Points(i)
134                     .UserSkills(i) = .UserSkills(i) + Points(i)
                    
                        'Client should prevent this, but just in case...
136                     If .UserSkills(i) > 100 Then
138                         .SkillPts = .SkillPts + .UserSkills(i) - 100
140                         .UserSkills(i) = 100
                        End If
                    
142                     Call CheckEluSkill(UserIndex, i, True)
                    End If

144             Next i

            End With
        End With

        '<EhFooter>
        Exit Sub

HandleModifySkills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleModifySkills " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTrain_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim SpawnedNpc As Integer

            Dim PetIndex   As Byte
        
102         PetIndex = Reader.ReadInt()
        
104         If .flags.TargetNPC = 0 Then Exit Sub
        
106         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
108         If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
110             If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                    'Create the creature
112                 SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
114                 If SpawnedNpc > 0 Then
116                     Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
118                     Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                    End If
                End If

            Else
120             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite))
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleTrain_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTrain " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCommerceBuy_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Slot   As Byte

            Dim Amount As Integer
            
            Dim SelectedPrice As Byte
            
102         Slot = Reader.ReadInt()
104         Amount = Reader.ReadInt()
              SelectedPrice = Reader.ReadInt8()
              
            'Dead people can't commerce...
106         If .flags.Muerto = 1 Then
108             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            '¿El target es un NPC valido?
110         If .flags.TargetNPC < 1 Then Exit Sub
            
            '¿El NPC puede comerciar?
112         If Npclist(.flags.TargetNPC).Comercia = 0 Then
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite))

                Exit Sub

            End If
        
            'Only if in commerce mode....
116         If Not .flags.Comerciando Then
118             Call WriteConsoleMsg(UserIndex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'User compra el item
120         Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount, SelectedPrice)
        End With

        '<EhFooter>
        Exit Sub

HandleCommerceBuy_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCommerceBuy " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBankExtractItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Slot     As Byte

            Dim Amount   As Integer
        
            Dim TypeBank As E_BANK
        
102         Slot = Reader.ReadInt()
104         Amount = Reader.ReadInt()
106         TypeBank = Reader.ReadInt()
        
108         If Slot <= 0 Then Exit Sub
        
            'Dead people can't commerce
110         If .flags.Muerto = 1 Then
112             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            '¿El target es un NPC valido?
114         If .flags.TargetNPC < 1 Then Exit Sub
        
            '¿Es el banquero?
116         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then

                Exit Sub

            End If
        
118         If .flags.SlotEvent > 0 Then
120             If Events(.flags.SlotEvent).ChangeClass > 0 Then
122                 Call WriteConsoleMsg(UserIndex, "En este tipo de eventos no es posible retirar/depositar objetos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
            End If
        
124         Select Case TypeBank

                Case E_BANK.e_User
126                 Call UserRetiraItem(UserIndex, Slot, Amount)

128             Case E_BANK.e_Account
130                 Call UserRetiraItem_Account(UserIndex, Slot, Amount)
            End Select
        
        End With

        '<EhFooter>
        Exit Sub

HandleBankExtractItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBankExtractItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCommerceSell_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Slot   As Byte

            Dim Amount As Integer
            
            Dim SelectedPrice As Byte
            
102         Slot = Reader.ReadInt()
104         Amount = Reader.ReadInt()
              SelectedPrice = Reader.ReadInt8
              
            'Dead people can't commerce...
106         If .flags.Muerto = 1 Then
108             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            '¿El target es un NPC valido?
110         If .flags.TargetNPC < 1 Then Exit Sub
        
            '¿El NPC puede comerciar?
112         If Npclist(.flags.TargetNPC).Comercia = 0 Then
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite))

                Exit Sub

            End If
        
              'User compra el item del slot
116         Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount, SelectedPrice)
        End With

        '<EhFooter>
        Exit Sub

HandleCommerceSell_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCommerceSell " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBankDeposit_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Slot     As Byte

            Dim Amount   As Integer
        
            Dim TypeBank As E_BANK
        
102         Slot = Reader.ReadInt()
104         Amount = Reader.ReadInt()
106         TypeBank = Reader.ReadInt()
        
108         If Slot <= 0 Or Amount <= 0 Then Exit Sub
        
            'Dead people can't commerce...
110         If .flags.Muerto = 1 Then
112             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            '¿El target es un NPC valido?
114         If .flags.TargetNPC < 1 Then Exit Sub
        
            '¿El NPC puede comerciar?
116         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then

                Exit Sub

            End If
        
118         If .flags.SlotEvent > 0 Then
120             If Events(.flags.SlotEvent).ChangeClass > 0 Then
122                 Call WriteConsoleMsg(UserIndex, "En este tipo de eventos no es posible retirar/depositar objetos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
            End If

124         Select Case TypeBank

                Case E_BANK.e_User
126                 Call UserDepositaItem(UserIndex, Slot, Amount)

128             Case E_BANK.e_Account
130                 Call UserDepositaItem_Account(UserIndex, Slot, Amount)
            End Select
        
        End With

        '<EhFooter>
        Exit Sub

HandleBankDeposit_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBankDeposit " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleMoveSpell_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        Dim dir As Integer
        
        Dim SlotOld As Byte
        Dim SlotNew As Byte
    
100     SlotOld = Reader.ReadInt
102     SlotNew = Reader.ReadInt
        
104     Call ChangeSlotSpell(UserIndex, SlotOld, SlotNew)

        '<EhFooter>
        Exit Sub

HandleMoveSpell_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleMoveSpell " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleMoveBank_Err
        '</EhHeader>

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 06/14/09
        '
        '***************************************************
        
        Dim dir      As Integer

        Dim Slot     As Byte

        Dim TempItem As Obj
        
100     If Reader.ReadBool() Then
102         dir = 1
        Else
104         dir = -1
        End If
        
106     Slot = Reader.ReadInt()

108     With UserList(UserIndex)
110         TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
112         TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
114         If dir = 1 Then 'Mover arriba
116             .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
118             .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
120             .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
            Else 'mover abajo
122             .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
124             .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
126             .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
            End If

        End With
    
128     Call UpdateBanUserInv(True, UserIndex, 0)
130     Call UpdateVentanaBanco(UserIndex)

        '<EhFooter>
        Exit Sub

HandleMoveBank_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleMoveBank " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUserCommerceOffer_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 24/11/2009
        '24/11/2009: ZaMa - Nuevo sistema de comercio
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Amount    As Long

            Dim Slot      As Byte

            Dim tUser     As Integer

            Dim OfferSlot As Byte

            Dim ObjIndex  As Integer
        
102         Slot = Reader.ReadInt()
104         Amount = Reader.ReadInt()
106         OfferSlot = Reader.ReadInt()
        
            'Get the other player
108         tUser = .ComUsu.DestUsu
        
            ' If he's already confirmed his offer, but now tries to change it, then he's cheating
110         If UserList(UserIndex).ComUsu.Confirmo = True Then
            
                ' Finish the trade
112             Call FinComerciarUsu(UserIndex)
        
114             If tUser <= 0 Or tUser > MaxUsers Then
116                 Call FinComerciarUsu(tUser)
118                 Call Protocol.FlushBuffer(tUser)
                End If
        
                Exit Sub

            End If
        
            'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
120         If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO And Slot <> FLAGELDHIR) Then Exit Sub
        
            'If OfferSlot is invalid, then ignore it.
122         If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 2 Then Exit Sub
        
            ' Can be negative if substracted from the offer, but never 0.
124         If Amount = 0 Then Exit Sub
        
            'Has he got enough??
126         If Slot = FLAGORO Then

                ' Can't offer more than he has
128             If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
130                 Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If
            
132             If Amount < 0 Then
134                 If Abs(Amount) > .ComUsu.GoldAmount Then
136                     Amount = .ComUsu.GoldAmount * (-1)
                    End If
                End If

138         ElseIf Slot = FLAGELDHIR Then

                ' Can't offer more than he has
140             If Amount > .Stats.Eldhir - .ComUsu.EldhirAmount Then
142                 Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de Eldhir para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If
            
144             If Amount < 0 Then
146                 If Abs(Amount) > .ComUsu.EldhirAmount Then
148                     Amount = .ComUsu.EldhirAmount * (-1)
                    End If
                End If

            Else

                'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
150             If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex

                ' Can't offer more than he has
152             If Not HasEnoughItems(UserIndex, ObjIndex, TotalOfferItems(ObjIndex, UserIndex) + Amount) Then
                
154                 Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If
            
156             If Amount < 0 Then
158                 If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
160                     Amount = .ComUsu.cant(OfferSlot) * (-1)
                    End If
                End If
        
162             If ItemNewbie(ObjIndex) Then
164                 Call WriteCancelOfferItem(UserIndex, OfferSlot)

                    Exit Sub

                End If
            
                'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
166             If .flags.Navegando = 1 Then
168                 If .Invent.BarcoSlot = Slot Then
170                     Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)

                        Exit Sub

                    End If
                End If
            
172             If .Invent.MochilaEqpSlot > 0 Then
174                 If .Invent.MochilaEqpSlot = Slot Then
176                     Call WriteCommerceChat(UserIndex, "No puedes vender tu mochila mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)

                        Exit Sub

                    End If
                End If
            
178             If ObjData(ObjIndex).OBJType = otGemaTelep Then
180                 Call WriteCommerceChat(UserIndex, "No puedes vender los scrolls de viajes.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If
            
182             If Not EsGmPriv(UserIndex) Then
184                 If ObjData(ObjIndex).NoNada = 1 Then
186                     Call WriteCommerceChat(UserIndex, "No puedes realizar ninguna acción con este objeto. ¡Podría ser de uso personal!", FontTypeNames.FONTTYPE_TALK)
    
                        Exit Sub
    
                    End If
                End If
            End If
        
188         Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO, Slot = FLAGELDHIR)
190         Call EnviarOferta(tUser, OfferSlot)
        End With

        '<EhFooter>
        Exit Sub

HandleUserCommerceOffer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUserCommerceOffer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleOnline_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i     As Long

        Dim Count As Long
    
100     With UserList(UserIndex)
        
            Dim ArmadasON As Long

            Dim CaosON    As Long

            Dim lstName   As String
            
            Dim lstCaos As String
            Dim lstArmada As String
            
            Dim ViewFaction As Boolean
            
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil Or PlayerType.ChaosCouncil)) <> 0 Then
                ViewFaction = True
            End If
            
102         For i = 1 To LastUser

104             If Len(UserList(i).Account.Email) > 0 Then

                    'If UserList(i).flags.Privilegios And (PlayerType.User ) Then
106                     If UserList(i).Faction.Status = r_Caos Then
108                         CaosON = CaosON + 1

                            If ViewFaction Then
                                lstCaos = lstCaos & UserList(i).Name & ", "
                            End If
110                     ElseIf UserList(i).Faction.Status = r_Armada Then
112                         ArmadasON = ArmadasON + 1

                            If ViewFaction Then
                                lstArmada = lstArmada & UserList(i).Name & ", "
                            End If
                        End If
                
114                 lstName = lstName & UserList(i).Name & ", "
116                 Count = Count + 1
                    
                    'End If
                End If

118         Next i
        
120         If Count > 0 Then
122             lstName = Left$(lstName, Len(lstName) - 2)
            End If
        
        
            If ViewFaction Then
                If ArmadasON > 0 Then
                    lstArmada = Left$(lstArmada, Len(lstArmada) - 2)
                End If
    
                If CaosON > 0 Then
                    lstCaos = Left$(lstCaos, Len(lstCaos) - 2)
                End If
            End If

            
124         Count = Count + UsersBot
126         Call WriteConsoleMsg(UserIndex, "Número de usuarios online: " & CStr(Count) & ". El Record de usuarios conectados simultaneamente fue de " & RECORDusuarios, FontTypeNames.FONTTYPE_INFO)
128         If EsGmPriv(UserIndex) Then
130             Call WriteConsoleMsg(UserIndex, "Nombres de los usuarios: " & lstName, FontTypeNames.FONTTYPE_INFO)
            End If
132         Call WriteConsoleMsg(UserIndex, "Usuarios de la facción <Legión Oscura>: " & CStr(CaosON) & ". " & lstCaos, FontTypeNames.FONTTYPE_INFORED)
134         Call WriteConsoleMsg(UserIndex, "Usuarios de la facción <Armada Real>: " & CStr(ArmadasON) & ". " & lstArmada, FontTypeNames.FONTTYPE_INFOGREEN)
        End With

        '<EhFooter>
        Exit Sub

HandleOnline_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleOnline " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Quit_AddNew(ByVal UserIndex As Integer, ByVal IsAccount As Boolean)
        '<EhHeader>
        On Error GoTo Quit_AddNew_Err
        '</EhHeader>

        Dim tUser As Integer
        
100     With UserList(UserIndex)

            'exit secure commerce
102         If .ComUsu.DestUsu > 0 Then
104             tUser = .ComUsu.DestUsu
            
106             If UserList(tUser).flags.UserLogged Then
108                 If UserList(tUser).ComUsu.DestUsu = UserIndex Then
110                     Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
112                     Call FinComerciarUsu(tUser)

                    End If

                End If
            
114             Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_TALK)
116             Call FinComerciarUsu(UserIndex)

            End If

118         .flags.DeslogeandoCuenta = IsAccount
120         Call Cerrar_Usuario(UserIndex)
    
        End With

        '<EhFooter>
        Exit Sub

Quit_AddNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.Quit_AddNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ)
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    '04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
    '***************************************************
    Dim tUser        As Integer
    Dim IsAccount As Boolean
        
    IsAccount = Reader.ReadBool
        
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)

            Exit Sub

        End If
        
        
        Quit_AddNew UserIndex, IsAccount
    
    End With

End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleMeditate_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/15/08 (NicoNZ)
        'Arreglé un bug que mandaba un index de la meditacion diferente
        'al que decia el server.
        '***************************************************
100     With UserList(UserIndex)
        
            'Si ya tiene el mana completo, no lo dejamos meditar.
101         If .Stats.MinMan = .Stats.MaxMan Then Exit Sub

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Can he meditate?
106         If .Stats.MaxMan = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub


            End If
            
            If .flags.TeleportInvoker > 0 Then
                Exit Sub
            End If
            
118         .flags.Meditando = Not .flags.Meditando

120         If .flags.Meditando Then
122             .Char.loops = INFINITE_LOOPS
                .Counters.TimerMeditar = 0
                .Counters.TiempoInicioMeditar = 0

124             If .MeditationSelected = 0 Then
126                 .Char.FX = UserFxMeditation(UserIndex)
                Else
128                 .Char.FX = Meditation(.MeditationSelected)

                End If
            
            Else
            
134             .Char.FX = 0

            End If

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, .Char.FX, .Pos.X, .Pos.Y))

        End With

        '<EhFooter>
        Exit Sub

HandleMeditate_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMeditate " & "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function UserFxMeditation(ByVal UserIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo UserFxMeditation_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

102         If .Stats.Elv < 15 Then
104             UserFxMeditation = FXIDs.FXMEDITARCHICO

106         ElseIf .Stats.Elv < 30 Then
108             UserFxMeditation = FXIDs.FXMEDITARMEDIANO

110         ElseIf .Stats.Elv < 45 Then
112             UserFxMeditation = FXIDs.FXMEDITARGRANDE ' Celeste Mediana
                
114         ElseIf .Stats.Elv < STAT_MAXELV Then
116             UserFxMeditation = FXIDs.FXMEDITARXGRANDE

            Else
118             UserFxMeditation = FXIDs.FXMEDITARXXXGRANDE
            End If

        End With
    
        '<EhFooter>
        Exit Function

UserFxMeditation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.UserFxMeditation " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleResucitate_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
            'Se asegura que el target es un npc
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
106         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
            'Make sure it's close enough
108         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
110             Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
112         Call RevivirUsuario(UserIndex)
              .Stats.MinHp = .Stats.MaxHp
              Call WriteUpdateHP(UserIndex)
114         Call WriteConsoleMsg(UserIndex, "¡¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleResucitate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleResucitate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)
        '<EhHeader>
        On Error GoTo HandleConsultation_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 01/05/2010
        'Habilita/Deshabilita el modo consulta.
        '01/05/2010: ZaMa - Agrego validaciones.
        '16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
        '***************************************************
    
        Dim UserConsulta As Integer
    
100     With UserList(UserIndex)
        
            ' Comando exclusivo para gms
102         If Not EsGm(UserIndex) Then Exit Sub
        
104         UserConsulta = .flags.TargetUser
        
            'Se asegura que el target es un usuario
106         If UserConsulta = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            ' No podes ponerte a vos mismo en modo consulta.
110         If UserConsulta = UserIndex Then Exit Sub
        
            ' No podes estra en consulta con otro gm
112         If EsGm(UserConsulta) Then
114             Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            Dim UserName As String

116         UserName = UserList(UserConsulta).Name
        
            ' Si ya estaba en consulta, termina la consulta
118         If UserList(UserConsulta).flags.EnConsulta Then
120             Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
122             Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
124             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Termino consulta con " & UserName)
126             UserList(UserConsulta).flags.EnConsulta = False
        
                ' Sino la inicia
            Else
128             Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
130             Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
132             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Inicio consulta con " & UserName)
            
134             With UserList(UserConsulta)
136                 .flags.EnConsulta = True
                
                    ' Pierde invi u ocu
138                 If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
140                     .flags.Oculto = 0
142                     .flags.Invisible = 0
144                     .Counters.TiempoOculto = 0
146                     .Counters.Invisibilidad = 0
                    
148                     If UserList(UserConsulta).flags.Navegando = 0 Then
150                         Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.charindex, False)
                        End If
                    End If

                End With

            End If
        
152         Call UsUaRiOs.SetConsulatMode(UserConsulta)
        End With

        '<EhFooter>
        Exit Sub

HandleConsultation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleConsultation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleHeal_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
            'Se asegura que el target es un npc
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
106         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
108         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
110             Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
112         .Stats.MinHp = .Stats.MaxHp
        
114         Call WriteUpdateHP(UserIndex)
        
116         Call WriteConsoleMsg(UserIndex, "¡¡Has sido curado!!", FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleHeal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleHeal " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestStats_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     Call SendUserStatsTxt(UserIndex, UserIndex)
        '<EhFooter>
        Exit Sub

HandleRequestStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleHelp_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     Call SendHelp(UserIndex)
        '<EhFooter>
        Exit Sub

HandleHelp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleHelp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCommerceStart_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i As Integer

100     With UserList(UserIndex)
        
            'Dead people can't commerce
102         If .flags.Muerto = 1 Then
104             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Is it already in commerce mode??
106         If .flags.Comerciando Then
108             Call WriteConsoleMsg(UserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
110         If .flags.SlotEvent > 0 Then Exit Sub
112         If .flags.SlotFast > 0 Then Exit Sub
114         If .flags.SlotReto > 0 Then Exit Sub
        
        
            'Validate target NPC
116         If .flags.TargetNPC > 0 Then

                'Does the NPC want to trade??
118             If Npclist(.flags.TargetNPC).Comercia = 0 Then
120                 If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
122                     Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                    End If
                
                    Exit Sub

                End If
            
124             If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 5 Then
126                 Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                'Start commerce....
128             Call IniciarComercioNPC(UserIndex)
                '[Alejo]
130         ElseIf .flags.TargetUser > 0 Then
        
                'User commerce...
                'Can he commerce??
132             If .flags.Privilegios And PlayerType.SemiDios Then
134                 Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)

                    Exit Sub

                End If
            
                'Is the other one dead??
136             If UserList(.flags.TargetUser).flags.Muerto = 1 Then
138                 Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                'Is it me??
140             If .flags.TargetUser = UserIndex Then
142                 Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                'Check distance
144             If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 5 Then
146                 Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                'Is he already trading?? is it with me or someone else??
148             If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
150                 Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                If .Stats.Elv < 4 Then
                    Call WriteConsoleMsg(UserIndex, "¡Entrena hasta Nivel 4 para usar este comando!", FontTypeNames.FONTTYPE_INFORED)
    
                    Exit Sub
    
                End If
            
                ' 133
152             If MapInfo(.Pos.Map).Pk Then
154                 Call WriteConsoleMsg(UserIndex, "No puedes comerciar en ZONA INSEGURA.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
156             If Not Interval_Commerce(UserIndex) Then
158                 Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar algunos segundos para enviar solicitud!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                'Initialize some variables...
160             .ComUsu.DestUsu = .flags.TargetUser
162             .ComUsu.DestNick = UserList(.flags.TargetUser).Name

164             For i = 1 To MAX_OFFER_SLOTS
166                 .ComUsu.cant(i) = 0
168                 .ComUsu.Objeto(i) = 0
170             Next i

172             .ComUsu.GoldAmount = 0
174             .ComUsu.EldhirAmount = 0
176             .ComUsu.Acepto = False
178             .ComUsu.Confirmo = False
            
                'Rutina para comerciar con otro usuario
180             Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
            Else
182             Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleCommerceStart_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCommerceStart " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBankStart_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        Dim TypeBank As E_BANK
    
100     TypeBank = Reader.ReadInt
    
102     With UserList(UserIndex)
        
            'Dead people can't commerce
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
108         If .flags.Comerciando Then
110             Call WriteConsoleMsg(UserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Validate target NPC
112         If .flags.TargetNPC > 0 Then
114             If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 5 Then
116                 Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                'If it's the banker....
118             If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then

120                 Select Case TypeBank

                        Case E_BANK.e_User, E_BANK.e_Account
122                         Call IniciarDeposito(UserIndex, TypeBank)
                        
124                     Case Else
                    End Select
                
                End If

            Else
128             Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleBankStart_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBankStart " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleShareNpc_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/04/2010
        'Shares owned npcs with other user
        '***************************************************
    
        Dim targetUserIndex  As Integer

        Dim SharingUserIndex As Integer
    
100     With UserList(UserIndex)
        
            ' Didn't target any user
102         targetUserIndex = .flags.TargetUser

104         If targetUserIndex = 0 Then Exit Sub
        
            ' Can't share with admins
106         If EsGm(targetUserIndex) Then
108             Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            ' Pk or Caos?
110         If Escriminal(UserIndex) Then

                ' Caos can only share with other caos
112             If esCaos(UserIndex) Then
114                 If Not esCaos(targetUserIndex) Then
116                     Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma facción!!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
                    ' Pks don't need to share with anyone
                Else

                    Exit Sub

                End If
        
                ' Ciuda or Army?
            Else

                ' Can't share
118             If Escriminal(targetUserIndex) Then
120                 Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            End If
        
            ' Already sharing with target
122         SharingUserIndex = .flags.ShareNpcWith

124         If SharingUserIndex = targetUserIndex Then Exit Sub
        
            ' Aviso al usuario anterior que dejo de compartir
126         If SharingUserIndex <> 0 Then
128             Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
130             Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            End If
        
132         .flags.ShareNpcWith = targetUserIndex
        
134         Call WriteConsoleMsg(targetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
136         Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(targetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleShareNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleShareNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleStopSharingNpc_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/04/2010
        'Stop Sharing owned npcs with other user
        '***************************************************
    
        Dim SharingUserIndex As Integer
    
100     With UserList(UserIndex)
        
102         SharingUserIndex = .flags.ShareNpcWith
        
104         If SharingUserIndex <> 0 Then
            
                ' Aviso al que compartia y al que le compartia.
106             Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            
110             .flags.ShareNpcWith = 0
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleStopSharingNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleStopSharingNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePartyMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim chat As String
        
102         chat = Reader.ReadString8()
        
104         If .GroupIndex = 0 Then
106             Call WriteConsoleMsg(UserIndex, "No conformas ninguna party", FontTypeNames.FONTTYPE_INFO)
            
            Else

108             If Interval_Message(UserIndex) Then
110                 If LenB(chat) <> 0 Then
112                     SendMessageGroup .GroupIndex, .Name, chat
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandlePartyMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePartyMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCouncilMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim chat As String
        
102         chat = Reader.ReadString8()
        
            Dim ValidChat As Boolean

104         ValidChat = True
        
106         If .flags.SlotEvent > 0 Then
108             If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then ValidChat = False
            End If
        
110         If LenB(chat) <> 0 And ValidChat Then
            
            
            
112             If .flags.Privilegios And PlayerType.RoyalCouncil Then
114                 Call SendData(SendTarget.ToConsejoYCaos, UserIndex, PrepareMessageConsoleMsg("(Privado Consejo) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
116             ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
118                 Call SendData(SendTarget.ToConsejoYCaos, UserIndex, PrepareMessageConsoleMsg("(Privado Concilio) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleCouncilMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCouncilMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeDescription_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '********

100     With UserList(UserIndex)
        
            Dim description As String
        
102         description = Reader.ReadString8()
        
104         If .Account.Premium > 1 Then
106             If .flags.Muerto = 1 Then
108                 Call WriteConsoleMsg(UserIndex, "No puedes cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
                Else

110                 If Not AsciiValidos(description) Then
112                     Call WriteConsoleMsg(UserIndex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
                    Else
114                     .Desc = Trim$(description)
116                     Call WriteConsoleMsg(UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If

            Else
118             Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior pueden cambiar la descripción de sus personajes. Consulta las promociones en /SHOP", FontTypeNames.FONTTYPE_INFO)
        
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeDescription_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeDescription " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePunishments_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 25/08/2009
        '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Name  As String

            Dim Count As Integer
        
102         Name = Reader.ReadString8()
        
104         If LenB(Name) <> 0 Then
106             If (InStrB(Name, "\") <> 0) Then
108                 Name = Replace(Name, "\", "")
                End If

110             If (InStrB(Name, "/") <> 0) Then
112                 Name = Replace(Name, "/", "")
                End If

114             If (InStrB(Name, ":") <> 0) Then
116                 Name = Replace(Name, ":", "")
                End If

118             If (InStrB(Name, "|") <> 0) Then
120                 Name = Replace(Name, "|", "")
                End If
            
122             If UCase$(Name) = UCase$(.Name) Then
124                 If FileExist(CharPath & Name & ".chr", vbNormal) Then
126                     Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))

128                     If Count = 0 Then
130                         Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                        Else

132                         While Count > 0

134                             Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
136                             Count = Count - 1

                            Wend

                        End If

                    Else
138                     Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else
            
140                 If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
142                     If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
144                         Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
                        Else

146                         If FileExist(CharPath & Name & ".chr", vbNormal) Then
148                             Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))

150                             If Count = 0 Then
152                                 Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                                Else

154                                 While Count > 0

156                                     Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
158                                     Count = Count - 1

                                    Wend

                                End If

                            Else
160                             Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandlePunishments_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePunishments " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGamble_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Amount  As Integer

            Dim TypeNpc As eNPCType
        
102         Amount = Reader.ReadInt()
        
            ' Dead?
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        
                'Validate target NPC
108         ElseIf .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        
                ' Validate Distance
112         ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
114             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
                ' Validate NpcType
116         ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            
                Dim TargetNpcType As eNPCType

118             TargetNpcType = Npclist(.flags.TargetNPC).NPCtype
            
                ' Normal npcs don't speak
120             If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
122                 Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                End If
            
                ' Validate amount
124         ElseIf Amount < 1 Then
126             Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
        
                ' Validate amount
128         ElseIf Amount > 50000 Then
130             Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 50000 Monedas de Oro.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
        
                ' Validate user gold
132         ElseIf .Stats.Gld < Amount Then
134             Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
        
            Else

136             If RandomNumber(1, 100) <= 47 Then
138                 .Stats.Gld = .Stats.Gld + Amount
140                 Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & CStr(Amount) & " Monedas de Oro.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                
142                 Apuestas.Perdidas = Apuestas.Perdidas + Amount
144                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
146                 .Stats.Gld = .Stats.Gld - Amount
148                 Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " Monedas de Oro.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                
150                 Apuestas.Ganancias = Apuestas.Ganancias + Amount
152                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
            
154             Apuestas.Jugadas = Apuestas.Jugadas + 1
            
156             Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
158             Call WriteUpdateGold(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleGamble_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGamble " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BankGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankGold(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleBankGold_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Amount  As Long

            Dim TypeGLD As Byte

            Dim Extract As Boolean
        
102         Amount = Reader.ReadInt()
104         TypeGLD = Reader.ReadInt()
106         Extract = Reader.ReadBool()
        
            'Dead people can't leave a faction.. they can't talk...
108         If .flags.Muerto = 1 Then
110             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Validate target NPC
112         If .flags.TargetNPC = 0 Then
114             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
116         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
118         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
120             Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
122         If .flags.SlotEvent > 0 Then
124             If Events(.flags.SlotEvent).ChangeClass > 0 Then
126                 Call WriteConsoleMsg(UserIndex, "En este tipo de eventos no es posible retirar/depositar objetos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If
        
128         Select Case TypeGLD

                Case 0 ' Monedas de Oro
                    
130                 If Extract Then
                          
132                     If (Amount > 0 And Amount <= .Account.Gld) Then

                            If Amount + .Stats.Gld < MAXORO Then
134                             .Account.Gld = .Account.Gld - Amount
136                             .Stats.Gld = .Stats.Gld + Amount
                                
138                             Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Gld & " Monedas de Oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)

                            End If

                        Else
140                         .Stats.Gld = .Stats.Gld + .Account.Gld
142                         .Account.Gld = 0

                        End If

                        If .Stats.Gld > MAXORO Then
                            .Stats.Gld = MAXORO

                        End If

                    Else

144                     If Amount > 0 And Amount <= .Stats.Gld Then
146                         .Account.Gld = .Account.Gld + Amount
148                         .Stats.Gld = .Stats.Gld - Amount
150                         Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Gld & " Monedas de Oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                                
                        Else
152                         .Account.Gld = .Account.Gld + .Stats.Gld
154                         .Stats.Gld = 0
156                         Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Gld & " Monedas de Oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
            
                        End If
                        
                        If .Account.Gld > MAXORO Then
                            .Account.Gld = MAXORO

                        End If

                    End If
                
158                 Call WriteUpdateGold(UserIndex)
            
160             Case 1 ' Monedas de Eldhir

162                 If Extract Then
164                     If Amount > 0 And Amount <= .Account.Eldhir Then
166                         .Account.Eldhir = .Account.Eldhir - Amount
168                         .Stats.Eldhir = .Stats.Eldhir + Amount
170                         Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Eldhir & " Monedas de Eldhir en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                        Else
172                         .Stats.Eldhir = .Stats.Eldhir + .Account.Eldhir
174                         .Account.Eldhir = 0

                        End If

                    Else

176                     If Amount > 0 And Amount <= .Stats.Eldhir Then
178                         .Account.Eldhir = .Account.Eldhir + Amount
180                         .Stats.Eldhir = .Stats.Eldhir - Amount
182                         Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Eldhir & " Monedas de Eldhir en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)

                        Else
184                         .Account.Eldhir = .Account.Eldhir + .Stats.Eldhir
186                         .Stats.Eldhir = 0
188                         Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Eldhir & " Monedas de Eldhir en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
            
                        End If
                
                    End If
                
190                 Call WriteUpdateDsp(UserIndex)
            
            End Select
        
192         Call WriteUpdateBankGold(UserIndex)
        
        End With

        '<EhFooter>
        Exit Sub

HandleBankGold_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBankGold " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDenounce_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 14/11/2010
        '14/11/2010: ZaMa - Now denounces can be desactivated.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Text As String

            Dim msg  As String
        
102         Text = Reader.ReadString8()
        
            Dim ValidChat As Boolean

104         ValidChat = True
        
106         If UCase$(Left$(Text, 11)) <> "[SEGURIDAD]" Then
108             If .flags.SlotEvent > 0 Then
110                 If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then ValidChat = False
                End If
            End If
        
112         If Len(Text) < 10 Then
114             Call WriteConsoleMsg(UserIndex, "Por favor, utiliza este comando para describir tu error de forma concreta. No solicites GMS, ni pongas cosas sin explicarlas de forma prolija. Queremos ayudarte rápido, ayudanos vos a nosotros", FontTypeNames.FONTTYPE_INFO)
116             ValidChat = False
            End If
        
118         If .flags.Silenciado = 0 And ValidChat And (.Counters.TimeDenounce = 0) Then
            
120             If UCase$(Left$(Text, 11)) = "[SEGURIDAD]" Then
122              '   .flags.ToleranceCheat = .flags.ToleranceCheat + 1

124                ' If .flags.ToleranceCheat >= 5 Then
126                    ' .flags.ToleranceCheat = 0
128                     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT] " & .Name & ": " & Text, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                          Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " IP: " & .Account.Sec.IP_Address & " Email: " & .Account.Email & " : " & Text)
                    'End If
                
130             ElseIf UCase$(Left$(Text, 15)) = "[FOTODENUNCIAS]" Then
132                 SendData SendTarget.ToGM, 0, PrepareMessageConsoleMsg(Text & ". Hecha por: " & .Name, FontTypeNames.FONTTYPE_INFO)
134                 .Counters.TimeDenounce = 20
                Else
136                 msg = LCase$(.Name) & " DENUNCIA: " & Text

138                 Call Denuncias.Push(msg, False)
        
140                 Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
142                 Call WriteConsoleMsg(UserIndex, "Denuncia enviada. Si quieres comunicarte mediante whatsapp y recibir una respuesta rápida ingresa a WWW.ARGENTUMGAME.COM", FontTypeNames.FONTTYPE_INFO)
144                 .Counters.TimeDenounce = 5
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleDenounce_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDenounce " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGMMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/08/07
        'Last Modification by: (liquid)
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

            Dim Priv    As Boolean
        
102        Message = Reader.ReadString8()
104        Priv = Reader.ReadBool()
        
106         If Not EsGm(UserIndex) Then Exit Sub
        
108         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje a Gms:" & Message)
        
110         If LenB(Message) <> 0 Then
112             Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_ADMIN))
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleGMMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGMMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleShowName_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And (PlayerType.Admin) Then
104             .ShowName = Not .ShowName 'Show / Hide the name
            
106             Call RefreshCharStatus(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleShowName_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleShowName " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleOnlineChaosLegion_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 28/05/2010
        '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And PlayerType.User Then Exit Sub
    
            Dim i    As Long

            Dim List As String

            Dim Priv As PlayerType

104         Priv = PlayerType.User Or PlayerType.SemiDios
        
            ' Solo dioses pueden ver otros dioses online
106         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
108             Priv = Priv Or PlayerType.Dios Or PlayerType.Admin
            End If
     
110         For i = 1 To LastUser

112             If UserList(i).ConnIDValida Then
114                 If UserList(i).Faction.Status = r_Caos Then
116                     If UserList(i).flags.Privilegios And Priv Then
118                         List = List & UserList(i).Name & ", "
                        End If
                    End If
                End If

120         Next i

        End With

122     If Len(List) > 0 Then
124         Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(List, Len(List) - 2), FontTypeNames.FONTTYPE_INFO)
        Else
126         Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

HandleOnlineChaosLegion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleOnlineChaosLegion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleServerTime_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
102         If Not EsGmPriv(UserIndex) Then Exit Sub
    
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Hora.")
        End With
    
106     Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
        '<EhFooter>
        Exit Sub

HandleServerTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleServerTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleWhere_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 18/11/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim miPos    As String
        
            Dim Guild As Boolean
        
            Dim GuildIndex As Integer
        
102         UserName = Reader.ReadString8()
104         Guild = Reader.ReadBool()
        
106         If Not EsGmPriv(UserIndex) Then Exit Sub

110         If Guild Then
112             GuildIndex = Guilds_SearchIndex(UCase$(UserName))
            
114             If GuildIndex = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "¡Clan inexistente!", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub
                End If
            
118             Call Guilds_PrepareOnline(UserIndex, GuildIndex)
            
            Else
120             tUser = NameIndex(UserName)
    
122             If tUser <= 0 Then
124                 If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    
126                     miPos = GetVar(CharPath & UserName & ".chr", "INIT", "POSITION")
128                     Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & " (Offline): " & ReadField(1, miPos, 45) & ", " & ReadField(2, miPos, 45) & ", " & ReadField(3, miPos, 45) & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
    
                Else
130                 Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
132         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/Donde " & UserName)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleWhere_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleWhere " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCreaturesInMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 30/07/06
        'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Map As Integer

            Dim i, j As Long

            Dim NPCcount1, NPCcount2 As Integer

            Dim NPCcant1() As Integer

            Dim NPCcant2() As Integer

            Dim List1()    As String

            Dim List2()    As String
        
102         Map = Reader.ReadInt()
        
104         If Not EsGmPriv(UserIndex) Then Exit Sub
        
106         If MapaValido(Map) Then

108             For i = 1 To LastNPC

                    'VB isn't lazzy, so we put more restrictive condition first to speed up the process
110                 If Npclist(i).Pos.Map = Map Then

                        '¿esta vivo?
112                     If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).flags.AIAlineacion = 2 Then
114                         If NPCcount1 = 0 Then
116                             ReDim List1(0) As String
118                             ReDim NPCcant1(0) As Integer
120                             NPCcount1 = 1
122                             List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
124                             NPCcant1(0) = 1
                            Else

126                             For j = 0 To NPCcount1 - 1

128                                 If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
130                                     List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
132                                     NPCcant1(j) = NPCcant1(j) + 1

                                        Exit For

                                    End If

134                             Next j

136                             If j = NPCcount1 Then
138                                 ReDim Preserve List1(0 To NPCcount1) As String
140                                 ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
142                                 NPCcount1 = NPCcount1 + 1
144                                 List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
146                                 NPCcant1(j) = 1
                                End If
                            End If

                        Else

148                         If NPCcount2 = 0 Then
150                             ReDim List2(0) As String
152                             ReDim NPCcant2(0) As Integer
154                             NPCcount2 = 1
156                             List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
158                             NPCcant2(0) = 1
                            Else

160                             For j = 0 To NPCcount2 - 1

162                                 If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
164                                     List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
166                                     NPCcant2(j) = NPCcant2(j) + 1

                                        Exit For

                                    End If

168                             Next j

170                             If j = NPCcount2 Then
172                                 ReDim Preserve List2(0 To NPCcount2) As String
174                                 ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
176                                 NPCcount2 = NPCcount2 + 1
178                                 List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
180                                 NPCcant2(j) = 1
                                End If
                            End If
                        End If
                    End If

182             Next i
            
184             Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)

186             If NPCcount1 = 0 Then
188                 Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
                Else

190                 For j = 0 To NPCcount1 - 1
192                     Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
194                 Next j

                End If

196             Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)

198             If NPCcount2 = 0 Then
200                 Call WriteConsoleMsg(UserIndex, "No hay más NPCS.", FontTypeNames.FONTTYPE_INFO)
                Else

202                 For j = 0 To NPCcount2 - 1
204                     Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
206                 Next j

                End If

208             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Numero enemigos en mapa " & Map)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleCreaturesInMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCreaturesInMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleWarpChar_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 26/03/2009
        '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim Map      As Integer

            Dim X        As Integer

            Dim Y        As Integer

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Map = Reader.ReadInt()
106         X = Reader.ReadInt()
108         Y = Reader.ReadInt()
        
110         If InStr(1, UserName, "+") Then
112             UserName = Replace(UserName, "+", " ")

            End If
        
114         If Not EsGm(UserIndex) Then Exit Sub
116         If Not MapaValido(Map) Then Exit Sub
              
118         If UCase$(UserName) = "YO" Then
120             tUser = UserIndex
                  
                ' @ Si no son DIOS, no pueden ir a ZONA INSEGURA.
                If Not EsGmDios(UserIndex) And MapInfo(Map).Pk = True Then Exit Sub
            Else
                ' @ Si no son DIOS NO PUEDEN TEPEAR USUARIOS
                If Not EsGmDios(UserIndex) Then Exit Sub
122             tUser = NameIndex(UserName)
            End If
        
124         If tUser <= 0 Then
126             If (EsDios(UserName) Or EsAdmin(UserName) And Not EsAdmin(.Name)) Then
128                 Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                Else

130                 If InMapBounds(Map, X, Y) Then
132                     If PersonajeExiste(UserName) Then
138                         Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Position", Map & "-" & X & "-" & Y)
140                         Call WriteConsoleMsg(UserIndex, "Usuario offline. Se ha modificado su posición.", FontTypeNames.FONTTYPE_INFO)
        
                        Else
142                         Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If
                    
            Else
                
                If Not EsGm(tUser) Then
146                 If Not CanUserTelep(Map, tUser) Then Exit Sub
                End If
                
150             If InMapBounds(Map, X, Y) Then
152                 If MapData(Map, X, Y).TileExit.Map = 0 Then
                        If UserList(tUser).PosAnt.Map <> Map Then
164                         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)

                        End If
                            
154                     UserList(tUser).PosAnt.Map = UserList(tUser).Pos.Map
156                     UserList(tUser).PosAnt.X = UserList(tUser).Pos.X
158                     UserList(tUser).PosAnt.Y = UserList(tUser).Pos.Y
                                    
160                     Call FindLegalPos(tUser, Map, X, Y)

                        If Map <> 0 And X <> 0 And Y <> 0 Then
162                         Call WarpUserChar(tUser, Map, X, Y, True, True)

                        End If

                    End If

                End If

            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleWarpChar_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWarpChar " & "at line " & Erl

        '</EhFooter>
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSilence_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If Not EsGmPriv(UserIndex) Then Exit Sub
        
106         tUser = NameIndex(UserName)
        
108         If tUser <= 0 Then
110             Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

112             If UserList(tUser).flags.Silenciado = 0 Then
114                 UserList(tUser).flags.Silenciado = 1
116                 Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
118                 Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias  y mensajes serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
120                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/silenciar " & UserList(tUser).Name)
                
                    'Flush the other user's buffer
122                 Call FlushBuffer(tUser)
                Else
124                 UserList(tUser).flags.Silenciado = 0
126                 Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
128                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleSilence_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSilence " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleGoToChar_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 26/03/2009
        '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim X        As Integer

            Dim Y        As Integer
        
            Dim Rank     As PlayerType
        
102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
        
106         If Not EsGmDios(UserIndex) Then Exit Sub ' Comando único para Gm's
            If Not EsGmPriv(UserIndex) And EsAdmin(UserName) Then Exit Sub
              
              
108         If tUser <= 0 Then
                   
112             Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                  If Not EsGmDios(UserIndex) And MapInfo(UserList(tUser).Pos.Map).Pk Then Exit Sub
                
                 
118             X = UserList(tUser).Pos.X
120             Y = UserList(tUser).Pos.Y + 1
122             Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
124             Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
126             If .flags.AdminInvisible = 0 Then
128                 Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
130                 Call FlushBuffer(tUser)

                End If
                    
132             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y & " (" & MapInfo(UserList(tUser).Pos.Map).Name & ")")

            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleGoToChar_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGoToChar " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleInvisible_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
102         If Not EsGm(UserIndex) Then Exit Sub
        
104         Call DoAdminInvisible(UserIndex)
106         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/INVISIBLE")
        End With

        '<EhFooter>
        Exit Sub

HandleInvisible_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleInvisible " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGMPanel_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmDios(UserIndex) Then Exit Sub
        
104         Call WriteShowGMPanelForm(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleGMPanel_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGMPanel " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestUserList_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/09/07
        'Last modified by: Lucas Tavolaro Ortiz (Tavo)
        'I haven`t found a solution to split, so i make an array of names
        '***************************************************
        Dim i       As Long

        Dim names() As String

        Dim Count   As Long
    
100     With UserList(UserIndex)
        
            If Not EsGmDios(UserIndex) Then Exit Sub
        
104         ReDim names(1 To LastUser) As String
106         Count = 1
        
108         For i = 1 To LastUser

110             If (LenB(UserList(i).Name) <> 0) Then
112                 If UserList(i).flags.Privilegios And PlayerType.User Then
114                     names(Count) = UserList(i).Name
116                     Count = Count + 1
                    End If
                End If

118         Next i
        
120         If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
        End With

        '<EhFooter>
        Exit Sub

HandleRequestUserList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestUserList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleJail_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 07/06/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim Reason   As String

            Dim jailTime As Byte

            Dim Count    As Byte

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         jailTime = Reader.ReadInt()
        
108         If InStr(1, UserName, "+") Then
110             UserName = Replace(UserName, "+", " ")
            End If
        
            If Not EsGmDios(UserIndex) Then Exit Sub
        
            '/carcel nick@motivo@<tiempo>
114             If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
                Else
118                 tUser = NameIndex(UserName)
                
120                 If tUser <= 0 Then
122                     If (EsDios(UserName) Or EsAdmin(UserName)) Then
124                         Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                        Else

126                         If FileExist(CharPath & UserName & ".chr", vbNormal) Then
128                             Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
130                             Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
132                             Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)
134                             Call WriteVar(CharPath & UserName & ".chr", "COUNTERS", "Pena", jailTime)
136                             Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", CStr(Prision.Map & "-" & Prision.X & "-" & Prision.Y))
                            End If
                        
138                         Call WriteConsoleMsg(UserIndex, "El usuario ha sido enviado a la carcel estando OFFLINE.", FontTypeNames.FONTTYPE_INFO)
                        End If

                    Else

140                     If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
142                         Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
144                     ElseIf jailTime > 60 Then
146                         Call WriteConsoleMsg(UserIndex, "No puedés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                        Else

148                         If (InStrB(UserName, "\") <> 0) Then
150                             UserName = Replace(UserName, "\", "")
                            End If

152                         If (InStrB(UserName, "/") <> 0) Then
154                             UserName = Replace(UserName, "/", "")
                            End If
                        
156                         If FileExist(CharPath & UserName & ".chr", vbNormal) Then
158                             Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
160                             Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
162                             Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)
                            End If
                        
164                         Call Encarcelar(tUser, jailTime, .Name)
166                         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, " encarceló a " & UserName)
                        End If
                    End If
                End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleJail_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleJail " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleKillNPC_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/22/08 (NicoNZ)
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGm(UserIndex) Then Exit Sub
        
            Dim tNpc   As Integer

            Dim auxNPC As Npc
        
104         tNpc = .flags.TargetNPC
        
106         If tNpc > 0 Then
108             If isNPCResucitador(tNpc) Then
110                 Call DeleteAreaResuTheNpc(tNpc)
                End If
        
112             Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNpc).Name, FontTypeNames.FONTTYPE_INFO)
            
114             auxNPC = Npclist(tNpc)
116             Call QuitarNPC(tNpc)
118             Call RespawnNpc(auxNPC)
            
120             .flags.TargetNPC = 0
            Else
122             Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleKillNPC_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleKillNPC " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleWarnUser_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/26/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim Reason   As String

            Dim Privs    As PlayerType

            Dim Count    As Byte

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
106         If EsGmDios(UserIndex) Then
108             If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
                Else
112                 Privs = UserDarPrivilegioLevel(UserName)
                
114                 If Not Privs And PlayerType.User Then
116                     Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                    Else

118                     If (InStrB(UserName, "\") <> 0) Then
120                         UserName = Replace(UserName, "\", "")
                        End If

122                     If (InStrB(UserName, "/") <> 0) Then
124                         UserName = Replace(UserName, "/", "")
                        End If
                    
126                     If FileExist(CharPath & UserName & ".chr", vbNormal) Then
128                         Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
130                         Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
132                         Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & Time)
                        
134                         tUser = NameIndex(UserName)
                        
136                         If tUser > 0 Then
138                             Call Encarcelar(tUser, 5)
                            Else
140                             Call WriteVar(CharPath & UserName & ".chr", "COUNTERS", "Pena", "5")
142                             Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", Prision.Map & "-" & Prision.X & "-" & Prision.Y)
                            End If
                        
144                         Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
146                         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, " advirtio a " & UserName)
                        End If
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleWarnUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleWarnUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestCharInfo_Err
        '</EhHeader>

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        'Last Modification by: (liquid).. alto bug zapallo..
        '***************************************************

100     With UserList(UserIndex)
                
            Dim TargetName  As String

            Dim TargetIndex As Integer
        
102         TargetName = Replace$(Reader.ReadString8(), "+", " ")
104         TargetIndex = NameIndex(TargetName)
        
106         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

                'is the player offline?
108             If TargetIndex <= 0 Then

                    'don't allow to retrieve administrator's info
110                 If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
112                     Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
                          
                          If EsGmPriv(UserIndex) Then
114                         Call SendUserStatsTxtOFF(UserIndex, TargetName)
                          End If
                    End If

                Else

118                   Call SendUserStatsTxt(UserIndex, TargetIndex)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRequestCharInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestCharInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestCharInventory_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/06/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName         As String

            Dim tUser            As Integer
        
            Dim UserIsAdmin      As Boolean

            Dim OtherUserIsAdmin As Boolean
        
102         UserName = Reader.ReadString8()
        
104         UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
106         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
108             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/INV " & UserName)
            
110             tUser = NameIndex(UserName)
112             OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
114             tUser = NameIndex(UserName)
116             OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
118             If tUser <= 0 Then
120                 If UserIsAdmin Or Not OtherUserIsAdmin Then
122                     Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                    
124                     Call SendUserInvTxtFromChar(UserIndex, UserName)
                    Else
126                     Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

128                 If UserIsAdmin Or Not OtherUserIsAdmin Then
130                     Call SendUserInvTxt(UserIndex, tUser)
                    Else
132                     Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRequestCharInventory_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestCharInventory " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequestCharBank_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/06/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName         As String

            Dim tUser            As Integer
        
            Dim UserIsAdmin      As Boolean

            Dim OtherUserIsAdmin As Boolean
        
            Dim TypeBank         As E_BANK
        
102         UserName = Reader.ReadString8()
104         TypeBank = Reader.ReadInt()
        
106         UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin)) <> 0
        
108         If UserIsAdmin Then
110             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BOV " & UserName)
            
112             tUser = NameIndex(UserName)
114             OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
116             tUser = NameIndex(UserName)
118             OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
120             If tUser <= 0 Then
122                 If UserIsAdmin Or Not OtherUserIsAdmin Then
124                     Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
126                     Select Case TypeBank

                            Case E_BANK.e_User
128                             Call SendUserBovedaTxtFromChar(UserIndex, UserName)

130                         Case E_BANK.e_Account
132                             Call SendUserBovedaTxtFromChar_Account(UserIndex, UserName)
                        End Select
                    
                    Else
134                     Call WriteConsoleMsg(UserIndex, "No puedes ver la bóveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

136                 If UserIsAdmin Or Not OtherUserIsAdmin Then

138                     Select Case TypeBank

                            Case E_BANK.e_User
140                             Call SendUserBovedaTxt(UserIndex, tUser)

142                         Case E_BANK.e_Account
144                             Call SendUserBovedaTxt_Account(UserIndex, tUser)
                        End Select
                    
                    Else
146                     Call WriteConsoleMsg(UserIndex, "No puedes ver la bóveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRequestCharBank_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequestCharBank " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleReviveChar_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 11/03/2010
        '11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
106             If UCase$(UserName) <> "YO" Then
108                 tUser = NameIndex(UserName)
                Else
110                 tUser = UserIndex
                End If
            
112             If tUser <= 0 Then
114                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

116                 With UserList(tUser)
                         If MapInfo(.Pos.Map).Pk Then Exit Sub
                         
                        'If dead, show him alive (naked).
118                     If .flags.Muerto = 1 Then
120                         .flags.Muerto = 0
                        
122                         If .flags.Navegando = 1 Then
124                             Call ToggleBoatBody(tUser)
                            Else
126                             Call DarCuerpoDesnudo(tUser)
                            End If
                        
128                         If .flags.Traveling = 1 Then
130                             Call EndTravel(tUser, True)
                            End If
                        
132                         Call ChangeUserChar(tUser, .Char.Body, .OrigChar.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                        
134                         Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                        Else
136                         Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    
138                     .Stats.MinHp = .Stats.MaxHp
                    
140                     If .flags.Traveling = 1 Then
142                         Call EndTravel(tUser, True)
                        End If
                    
                    End With
                
144                 Call WriteUpdateHP(tUser)
                
146                 Call FlushBuffer(tUser)
                
148                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Resucito a " & UserName)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleReviveChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleReviveChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleOnlineGM_Err
        '</EhHeader>

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 12/28/06
        '
        '***************************************************
        Dim i    As Long

        Dim List As String

        Dim Priv As PlayerType
    
100     With UserList(UserIndex)
        
102         If Not EsGm(UserIndex) Then Exit Sub

108         For i = 1 To LastUser

110             If UserList(i).flags.UserLogged Then
                      If EsGm(i) And Not EsGmPriv(i) Then
112                     List = List & UserList(i).Name & ", "
                      End If
                End If

114         Next i
        
116         If LenB(List) <> 0 Then
118             List = Left$(List, Len(List) - 2)
120             Call WriteConsoleMsg(UserIndex, List & ".", FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleOnlineGM_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleOnlineGM " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleOnlineMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 23/03/2009
        '23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
        '***************************************************
100     With UserList(UserIndex)
        
            Dim Map As Integer

102         Map = Reader.ReadInt
        
104         If Not EsGmPriv(UserIndex) Then Exit Sub
        
            Dim LoopC As Long

            Dim List  As String

            Dim Priv  As PlayerType
        
106         For LoopC = 1 To LastUser

108             If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
110                 List = List & UserList(LoopC).Name & ", "
                End If

112         Next LoopC
        
114         If Len(List) > 2 Then List = Left$(List, Len(List) - 2)
        
116         Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & List, FontTypeNames.FONTTYPE_INFO)
118         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/ONLINEMAP " & Map)
        End With

        '<EhFooter>
        Exit Sub

HandleOnlineMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleOnlineMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleForgive_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/06/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName    As String

            Dim tUser       As Integer

            Dim ResetArmada As Boolean
        
102         UserName = Reader.ReadString8()
104         ResetArmada = Reader.ReadBool()
            
            
            If ResetArmada Then
                If Not (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil)) <> 0 Then Exit Sub
            Else
                If Not (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil)) <> 0 Then Exit Sub
            End If
            
106         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil Or PlayerType.ChaosCouncil)) <> 0 Then
108             tUser = NameIndex(UserName)
            
110             If tUser > 0 Then
112                 If UserList(tUser).Faction.Status <> r_None Then
114                     Call WriteConsoleMsg(UserIndex, "El personaje ya pertenece a alguna facción.", FontTypeNames.FONTTYPE_INFO)
                    Else
116                     Call WriteConsoleMsg(UserIndex, "Has perdonado al personaje " & UserList(tUser).Name & ". " & IIf((ResetArmada = True), "Se reiniciaron Frags de Ciudadanos: EX VALOR: " & UserList(tUser).Faction.FragsCiu, vbNullString), FontTypeNames.FONTTYPE_INFOGREEN)
                    
118                     If ResetArmada Then UserList(tUser).Faction.FragsCiu = 0

120                     Call Faction_RemoveUser(tUser)
                    
122                     Call LogPerdones("El GM " & .Name & " ha perdonado al personaje " & UserList(tUser).Name & ".")
                    End If
                
                Else
124                 Call WriteConsoleMsg(UserIndex, "El personaje esta offline", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleForgive_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleForgive " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleKick_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/06/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim Rank     As Integer

            Dim IsAdmin  As Boolean
        
102         Rank = PlayerType.Admin Or PlayerType.Dios
        
104         UserName = Reader.ReadString8()
106         IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
108         If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
110             tUser = NameIndex(UserName)
            
112             If tUser <= 0 Then
114                 If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
116                     Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                    Else
118                     Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

120                 If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
122                     Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                    Else
124                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echó a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
126                     Call WriteDisconnect(tUser)
128                     Call FlushBuffer(tUser)
                        
130                     Call CloseSocket(tUser)
132                     Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Echó a " & UserName)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleKick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleKick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleExecute_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/06/2010
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
106             tUser = NameIndex(UserName)
            
108             If tUser > 0 Then
110                 If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
112                     Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                    Else

114                     If UserList(tUser).flags.Desafiando = 0 And _
                           UserList(tUser).flags.SlotReto = 0 And _
                           UserList(tUser).flags.SlotFast = 0 And _
                           UserList(tUser).flags.SlotEvent = 0 Then
                        
116                         Call UserDie(tUser)
118                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
120                         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, " ejecuto a " & UserName)
                    
                        Else
122                         Call WriteConsoleMsg(UserIndex, "El usuario no puede ser ejecutado en este momento.", FontTypeNames.FONTTYPE_INFO)
                    
                        End If
                    End If

                Else

124                 If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
126                     Call WriteConsoleMsg(UserIndex, "No está online.", FontTypeNames.FONTTYPE_INFO)
                    Else
128                     Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleExecute_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleExecute " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBanChar_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
        
        Dim UserName As String

        Dim Reason   As String
        
        Dim Tipo     As Byte
        
        Dim DataDay  As String
        
100     UserName = Reader.ReadString8()
102     Reason = Reader.ReadString8()
104     Tipo = Reader.ReadInt()
106     DataDay = Reader.ReadString8()
        
        
         If Not EsGmDios(UserIndex) Then Exit Sub
        
110     Select Case Tipo

            Case 0 ' Baneo de personajes
112             Call BanCharacter(UserIndex, UserName, Reason, DataDay)
            
114         Case 1 ' Baneo de cuenta
116             Call BanCharacter_Account(UserIndex, UserName, Reason, DataDay)
        End Select
    
        '<EhFooter>
        Exit Sub

HandleBanChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBanChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUnbanChar_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName  As String

            Dim cantPenas As Byte
        
102         UserName = Reader.ReadString8()
        
              If EsGmDios(UserIndex) Then
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")
                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")
                End If
            
114             If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
116                 Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
                Else

118                 If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
120                     Call UnBan(UserName)
                
                        'penas
122                     cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
124                     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
126                     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & Time)
                
128                     Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/UNBAN a " & UserName)
130                     Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                    Else
132                     Call WriteConsoleMsg(UserIndex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleUnbanChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUnbanChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleNPCFollow_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGm(UserIndex) Then Exit Sub
        
104         If .flags.TargetNPC > 0 Then
106             Call DoFollow(.flags.TargetNPC, .Name)
108             Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
110             Npclist(.flags.TargetNPC).flags.Paralizado = 0
112             Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleNPCFollow_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleNPCFollow " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        Dim IsEvent  As Byte
        
        UserName = Reader.ReadString8()
        IsEvent = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If (EsDios(UserName) Or EsAdmin(UserName)) And Not EsAdmin(.Name) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.User)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.SemiDios)) <> 0 Then
                  
                    If Not IsEvent Then

                        ' Usuario participando en otro eventos
                        If Not CanUserTelep(.Pos.Map, tUser) Then
                            WriteConsoleMsg UserIndex, "El personaje no está disponible para ser sumoneado.", FontTypeNames.FONTTYPE_INFO
                            Exit Sub
        
                        End If

                        If MapInfo(.Pos.Map).Pk Then
                            WriteConsoleMsg UserIndex, "El personaje no está disponible para ser sumoneado.", FontTypeNames.FONTTYPE_INFO
                            Exit Sub
        
                        End If
        
                        UserList(tUser).PosAnt.Map = UserList(tUser).Pos.Map
                        UserList(tUser).PosAnt.X = UserList(tUser).Pos.X
                        UserList(tUser).PosAnt.Y = UserList(tUser).Pos.Y
                                
                        Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                    Else

                        If Not UserList(tUser).flags.SlotEvent > 0 Then Exit Sub    ' @ Si no está en evento no puede usar el comando, capaz tenga que refresh.

                    End If
                          
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSpawnListRequest_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And (PlayerType.User) Then Exit Sub
        
104         Call EnviarSpawnList(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleSpawnListRequest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSpawnListRequest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSpawnCreature_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Npc As Integer
            Dim NpcIndex As Integer
            
102         Npc = Reader.ReadInt()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then

                  If MapInfo(.Pos.Map).Pk Then Exit Sub
                  
106             If Npc > 0 And Npc <= UBound(Declaraciones.SpawnList()) Then
                    
                    NpcIndex = SpawnNpc(Declaraciones.SpawnList(Npc).NpcIndex, .Pos, True, False)

108             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Sumoneo " & Declaraciones.SpawnList(Npc).NpcName)

End If
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleSpawnCreature_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSpawnCreature " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleResetNPCInventory_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGm(UserIndex) Then Exit Sub
104         If .flags.TargetNPC = 0 Then Exit Sub
        
106         Call ResetNpcInv(.flags.TargetNPC)
108         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
        End With

        '<EhFooter>
        Exit Sub

HandleResetNPCInventory_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleResetNPCInventory " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCleanWorld_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(UserIndex)

102         If Not EsGm(UserIndex) Then Exit Sub
        
104         Call LimpiarMundo
            'CountDownLimpieza = 5
        End With

        '<EhFooter>
        Exit Sub

HandleCleanWorld_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCleanWorld " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleServerMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 28/05/2010
        '28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
106             If LenB(Message) <> 0 Then
108                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje Broadcast:" & Message)
110                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "» " & Message, FontTypeNames.FONTTYPE_RMSG, eMessageType.Admin))
                    ''''''''''''''''SOLO PARA EL TESTEO'''''''
                    ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                    'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleServerMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleServerMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleMapMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
106             If LenB(Message) <> 0 Then
                
                    Dim mapa As Integer

108                 mapa = .Pos.Map
                
110                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje a mapa " & mapa & ":" & Message)
112                 Call SendData(SendTarget.toMap, mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleMapMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleMapMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleNickToIP_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 07/06/2010
        'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
        '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim Priv     As PlayerType

            Dim IsAdmin  As Boolean
        
102         UserName = Reader.ReadString8()
        
104         If EsGmPriv(UserIndex) Then
106             tUser = NameIndex(UserName)
108             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "NICK2IP Solicito la IP de " & UserName)
            
110             IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0

112             If IsAdmin Then
114                 Priv = PlayerType.User Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
                Else
116                 Priv = PlayerType.User
                End If
            
118             If tUser > 0 Then
120                 If UserList(tUser).flags.Privilegios And Priv Then
122                     Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).IpAddress, FontTypeNames.FONTTYPE_INFO)

                        Dim IP    As String

                        Dim lista As String

                        Dim LoopC As Long

124                     IP = UserList(tUser).IpAddress

126                     For LoopC = 1 To LastUser

128                         If UserList(LoopC).IpAddress = IP Then
130                             If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
132                                 If UserList(LoopC).flags.Privilegios And Priv Then
134                                     lista = lista & UserList(LoopC).Name & ", "
                                    End If
                                End If
                            End If

136                     Next LoopC

138                     If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
140                     Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

142                 If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
144                     Call WriteConsoleMsg(UserIndex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleNickToIP_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleNickToIP " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleIPToNick_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim IP    As String

            Dim LoopC As Long

            Dim lista As String

            Dim Priv  As PlayerType
        
102         IP = Reader.ReadInt() & "."
104         IP = IP & Reader.ReadInt() & "."
106         IP = IP & Reader.ReadInt() & "."
108         IP = IP & Reader.ReadInt()
        
110         If Not EsGmPriv(UserIndex) Then Exit Sub
        
112         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "IP2NICK Solicito los Nicks de IP " & IP)
        
114         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
116             Priv = PlayerType.User Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
118             Priv = PlayerType.User
            End If

120         For LoopC = 1 To LastUser

122             If UserList(LoopC).IpAddress = IP Then
124                 If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
126                     If UserList(LoopC).flags.Privilegios And Priv Then
128                         lista = lista & UserList(LoopC).Name & ", "
                        End If
                    End If
                End If

130         Next LoopC
        
132         If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
134         Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleIPToNick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleIPToNick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTeleportCreate_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 22/03/2010
        '15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
        '22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim mapa  As Integer

            Dim X     As Byte

            Dim Y     As Byte

            Dim Radio As Byte
        
102         mapa = Reader.ReadInt()
104         X = Reader.ReadInt()
106         Y = Reader.ReadInt()
108         Radio = Reader.ReadInt()
        
110         Radio = MinimoInt(Radio, 6)
        
112         If Not EsGm(UserIndex) Then Exit Sub
  
120         If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then Exit Sub
        
122         If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
124         If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub


            If Not EsGmPriv(UserIndex) Then
                If MapInfo(.Pos.Map).Pk Then Exit Sub
    
                ' Crea con destino inseguro y es semi dios
                If Not EsGmDios(UserIndex) Then
                   If MapInfo(mapa).Pk Then Exit Sub
               End If
            End If
            
            
126         If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
128             Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
130         If MapData(mapa, X, Y).TileExit.Map > 0 Then
132             Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            Dim ET As Obj

134         ET.Amount = 1
            ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
136         ET.ObjIndex = TELEP_OBJ_INDEX 'TELEP_OBJ_INDEX + Radio
        
138         With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
140             .TileExit.Map = mapa
142             .TileExit.X = X
144             .TileExit.Y = Y
            End With
        
146         Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/CT " & mapa & "," & X & "," & Y & "," & Radio)
        End With

        '<EhFooter>
        Exit Sub

HandleTeleportCreate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTeleportCreate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTeleportDestroy_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(UserIndex)

            Dim mapa As Integer

            Dim X    As Byte

            Dim Y    As Byte
        
            '/dt
            
            If Not EsGm(UserIndex) Then Exit Sub
            
104         mapa = .flags.TargetMap
106         X = .flags.TargetX
108         Y = .flags.TargetY
        
110         If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
112         With MapData(mapa, X, Y)

114             If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
116             If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
118                 Call Logs_User(UserList(UserIndex).Name, eLog.eGm, eLogDescUser.eNone, "/DT: " & mapa & "," & X & "," & Y)
                
120                 Call EraseObj(.ObjInfo.Amount, mapa, X, Y)
                
122                 If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
124                     Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                    End If
                
126                 .TileExit.Map = 0
128                 .TileExit.X = 0
130                 .TileExit.Y = 0
                End If

            End With
        End With

        '<EhFooter>
        Exit Sub

HandleTeleportDestroy_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTeleportDestroy " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "EnableDenounces" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnableDenounces(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEnableDenounces_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        'Enables/Disables
        '***************************************************

100     With UserList(UserIndex)
        
102       If Not EsGmPriv(UserIndex) Then Exit Sub
        
            Dim Activado As Boolean

            Dim msg      As String
        
104         Activado = Not .flags.SendDenounces
106         .flags.SendDenounces = Activado
        
108         msg = "Denuncias por consola " & IIf(Activado, "activadas", "desactivadas") & "."
        
110         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, msg)
        
112         Call WriteConsoleMsg(UserIndex, msg, FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleEnableDenounces_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEnableDenounces " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ShowDenouncesList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowDenouncesList(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleShowDenouncesList_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And PlayerType.User Then Exit Sub
104         Call WriteShowDenounces(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleShowDenouncesList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleShowDenouncesList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HanldeForceMIDIToMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
            Dim midiID As Byte

            Dim mapa   As Integer
        
102         midiID = Reader.ReadInt
104         mapa = Reader.ReadInt
        
            'Solo dioses, admins y RMS
106         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then

                'Si el mapa no fue enviado tomo el actual
108             If Not InMapBounds(mapa, 50, 50) Then
110                 mapa = .Pos.Map
                End If
        
112             If midiID = 0 Then
                    'Ponemos el default del mapa
114                 'Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMusic(MapInfo(.Pos.Map).Music))
                Else
                    'Ponemos el pedido por el GM
116                 'Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMusic(midiID))
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

HanldeForceMIDIToMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HanldeForceMIDIToMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleForceWAVEToMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim waveID As Integer

            Dim mapa   As Integer

            Dim X      As Byte

            Dim Y      As Byte
        
102         waveID = Reader.ReadInt()
104         mapa = Reader.ReadInt()
106         X = Reader.ReadInt()
108         Y = Reader.ReadInt()
        
            'Solo dioses, admins y RMS
110         If EsGmDios(UserIndex) Then

                'Si el mapa no fue enviado tomo el actual
112             If Not InMapBounds(mapa, X, Y) Then
114                 mapa = .Pos.Map
116                 X = .Pos.X
118                 Y = .Pos.Y
                End If
            
                'Ponemos el pedido por el GM
120             'Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayEffect(waveID, X, Y))
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleForceWAVEToMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleForceWAVEToMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRoyalArmyMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8()
        
            'Solo dioses, admins, semis y RMS
104         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoyalCouncil) Then
106             Call SendData(SendTarget.ToCiudadanos, 0, PrepareMessageConsoleMsg("[Consejo de Banderbill] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_CONSEJOVesA))
        
            Else
108             If .Faction.Status = r_Armada Then
110                 Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("[Armada Real] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_INFOGREEN))
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRoyalArmyMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRoyalArmyMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChaosLegionMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8()
        
            'Solo dioses, admins, concilios
104         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.ChaosCouncil) Then
106             Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("[Concilio de las Sombras] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_EJECUCION))
            Else
            
108             If .Faction.Status = r_Caos Then
110                 Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("[Legión Oscura] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_INFORED))
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChaosLegionMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChaosLegionMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleTalkAsNPC_Err

        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8()
        
              ' Solo dioses, admins y RMS
104         If EsGmPriv(UserIndex) Then

                'Asegurarse haya un NPC seleccionado
106             If .flags.TargetNPC > 0 Then
108                 Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.charindex, vbWhite))
                Else
110                 Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleTalkAsNPC_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTalkAsNPC " & "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDestroyAllItemsInArea_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
            If Not EsGm(UserIndex) Then Exit Sub

            Dim X       As Long

            Dim Y       As Long

            Dim bIsExit As Boolean
        
104         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
106             For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

108                 If X > 0 And Y > 0 And X < 101 And Y < 101 Then
110                     If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
112                         bIsExit = MapData(.Pos.Map, X, Y).TileExit.Map > 0

114                         If ItemNoEsDeMapa(.Pos.Map, X, Y, bIsExit) Then
116                             Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                            End If
                        End If
                    End If

118             Next X
120         Next Y
        
122         Call Logs_User(UserList(UserIndex).Name, eLog.eGm, eLogDescUser.eNone, "/MASSDEST")
        End With

        '<EhFooter>
        Exit Sub

HandleDestroyAllItemsInArea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDestroyAllItemsInArea " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAcceptRoyalCouncilMember_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
                Else
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

114                 With UserList(tUser)

116                     If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
118                     If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
120                     Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                    End With

                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleAcceptRoyalCouncilMember_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAcceptRoyalCouncilMember " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAcceptChaosCouncilMember_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
                Else
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
114                 With UserList(tUser)

116                     If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
118                     If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

120                     Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                    End With

                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleAcceptChaosCouncilMember_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAcceptChaosCouncilMember " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleItemsInTheFloor_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios) Then Exit Sub
        
            Dim tobj  As Integer

            Dim lista As String

            Dim X     As Long

            Dim Y     As Long
        
104         For X = 5 To 95
106             For Y = 5 To 95
108                 tobj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex

110                 If tobj > 0 Then
112                     If ObjData(tobj).OBJType <> eOBJType.otArboles Then
114                         Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tobj).Name, FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If

116             Next Y
118         Next X

        End With

        '<EhFooter>
        Exit Sub

HandleItemsInTheFloor_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleItemsInTheFloor " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCouncilKick_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
106             tUser = NameIndex(UserName)

108             If tUser <= 0 Then
110                 If FileExist(CharPath & UserName & ".chr") Then
112                     Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
114                     Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
116                     Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                    Else
118                     Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

120                 With UserList(tUser)

122                     If .flags.Privilegios And PlayerType.RoyalCouncil Then
124                         Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
126                         .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
128                         Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
130                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                        End If
                    
132                     If .flags.Privilegios And PlayerType.ChaosCouncil Then
134                         Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
136                         .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
138                         Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
140                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                        End If

                    End With

                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleCouncilKick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCouncilKick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSetTrigger_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim tTrigger As Byte

            Dim tLog     As String

            Dim ObjIndex As Integer
        
102         tTrigger = Reader.ReadInt()
        
104         If Not EsGmDios(UserIndex) Then Exit Sub
        
106         If tTrigger >= 0 Then
108             If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura Then
110                 If tTrigger <> eTrigger.zonaOscura Then
112                     If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
                    
114                     ObjIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                    
116                     If ObjIndex > 0 Then
118                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectCreate(ObjIndex, ObjData(ObjIndex).GrhIndex, .Pos.X, .Pos.Y, vbNullString, 0, ObjData(ObjIndex).Sound))
                        End If
                    End If

                Else

120                 If tTrigger = eTrigger.zonaOscura Then
122                     If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
                    
124                     ObjIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                    
126                     If ObjIndex > 0 Then
128                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectDelete(.Pos.X, .Pos.Y))
                        End If
                    End If
                End If
            
130             MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
132             tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
134             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, tLog)
136             Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleSetTrigger_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSetTrigger " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAskTrigger_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 04/13/07
        '
        '***************************************************
        Dim tTrigger As Byte
    
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios) Then Exit Sub
        
104         tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
106         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
108         Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleAskTrigger_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAskTrigger " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBannedIPList_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmPriv(UserIndex) Then Exit Sub
        
            Dim lista As String

            Dim LoopC As Long
        
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BANIPLIST")
        
106         For LoopC = 1 To BanIps.Count
108             lista = lista & BanIps.Item(LoopC) & ", "
110         Next LoopC
        
112         If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
114         Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleBannedIPList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBannedIPList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBannedIPReload_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmPriv(UserIndex) Then Exit Sub
        
104         Call BanIpGuardar
106         Call BanIpCargar
        End With

        '<EhFooter>
        Exit Sub

HandleBannedIPReload_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBannedIPReload " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleBanIP_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 07/02/09
        'Agregado un CopyBuffer porque se producia un bucle
        'inifito al intentar banear una ip ya baneada. (NicoNZ)
        '07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim bannedIP As String

            Dim tUser    As Integer

            Dim Reason   As String

            Dim i        As Long
        
            ' Is it by ip??
102         If Reader.ReadBool() Then
104             bannedIP = Reader.ReadInt() & "."
106             bannedIP = bannedIP & Reader.ReadInt() & "."
108             bannedIP = bannedIP & Reader.ReadInt() & "."
110             bannedIP = bannedIP & Reader.ReadInt()
            Else
112             tUser = NameIndex(Reader.ReadString8())
            
114             If tUser > 0 Then bannedIP = UserList(tUser).IpAddress
            End If
        
116         Reason = Reader.ReadString8()
        
118         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
120             If LenB(bannedIP) > 0 Then
122                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BanIP " & bannedIP & " por " & Reason)
                
124                 If BanIpBuscar(bannedIP) > 0 Then
126                     Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Else
128                     Call BanIpAgrega(bannedIP)
130                     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                        'Find every player with that ip and ban him!
132                     For i = 1 To LastUser

134                         If UserList(i).ConnIDValida Then
136                             If UserList(i).IpAddress = bannedIP Then
138                                 Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & Reason)
                                End If
                            End If

140                     Next i

                    End If

142             ElseIf tUser <= 0 Then
144                 Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleBanIP_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleBanIP " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUnbanIP_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/30/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim bannedIP As String
        
102         bannedIP = Reader.ReadInt() & "."
104         bannedIP = bannedIP & Reader.ReadInt() & "."
106         bannedIP = bannedIP & Reader.ReadInt() & "."
108         bannedIP = bannedIP & Reader.ReadInt()
        
110         If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios) Then Exit Sub
        
112         If BanIpQuita(bannedIP) Then
114             Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            Else
116             Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleUnbanIP_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUnbanIP " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCreateItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
    
100     With UserList(UserIndex)

            Dim tobj As Integer

            Dim tStr As String

102         tobj = Reader.ReadInt()

104         If Not EsGmPriv(UserIndex) Then Exit Sub
            
            Dim mapa As Integer

            Dim X    As Byte

            Dim Y    As Byte
        
106         mapa = .Pos.Map
108         X = .Pos.X
110         Y = .Pos.Y
            
112         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/CI: " & tobj & " en mapa " & mapa & " (" & X & "," & Y & ")")
        
114         If MapData(mapa, X, Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
116         If MapData(mapa, X, Y - 1).TileExit.Map > 0 Then Exit Sub
        
118         If tobj < 1 Or tobj > NumObjDatas Then Exit Sub
        
            'Is the object not null?
120         If LenB(ObjData(tobj).Name) = 0 Then Exit Sub
                
                If Not EsGmPriv(UserIndex) Then
                
                    'Silla
                    'Trono
                    'Sillon
                    'Silla
                    
                    If tobj <> 882 And _
                        tobj <> 162 And _
                        tobj <> 168 And _
                        tobj <> 826 Then
                        
                        Exit Sub
                        
                    End If
162
168
826
                End If
                
            Dim Objeto As Obj
            
            
            'NoCrear = 1
            
122         Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN: FUERON CREADOS ***25*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
124         Objeto.Amount = 25
126         Objeto.ObjIndex = tobj
128         Call MakeObj(Objeto, mapa, X, Y - 1)

130         Call Logs_User(.Name, eGm, eNone, "/CI: [" & tobj & "]" & ObjData(tobj).Name & " en mapa " & mapa & " (" & X & "," & Y & ")")
        
        End With

        '<EhFooter>
        Exit Sub

HandleCreateItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCreateItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDestroyItems_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGm(UserIndex) Then Exit Sub
        
            Dim mapa As Integer

            Dim X    As Byte

            Dim Y    As Byte
        
104         mapa = .Pos.Map
106         X = .Pos.X
108         Y = .Pos.Y
        
            Dim ObjIndex As Integer

110         ObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
        
112         If ObjIndex = 0 Then Exit Sub
        
114         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/DEST " & ObjIndex & " en mapa " & mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(mapa, X, Y).ObjInfo.Amount)
        
116         If ObjData(ObjIndex).OBJType = eOBJType.otTeleport And MapData(mapa, X, Y).TileExit.Map > 0 Then
            
118             Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
120         Call EraseObj(10000, mapa, X, Y)
        End With

        '<EhFooter>
        Exit Sub

HandleDestroyItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDestroyItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChaosLegionKick_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil)) <> 0 Or .flags.PrivEspecial Then
            
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")
                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")
                End If

114             tUser = NameIndex(UserName)
            
116             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "ECHO DEL CAOS A: " & UserName)
    
118             If tUser > 0 Then
120                 If .Faction.Status > 0 Then
122                     Call mFacciones.Faction_RemoveUser(tUser)
124                     Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
126                     Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
128                     Call FlushBuffer(tUser)
                    Else
130                     Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

132                 If FileExist(CharPath & UserName & ".chr") Then
                
                        Dim Status As Byte

134                     Status = val(GetVar(CharPath & UserName & ".chr", "FACTION", "STATUS"))
                    
136                     If Status > 0 Then
138                         Call WriteVar(CharPath & UserName & ".chr", "FACTION", "STATUS", "0")
140                         Call WriteVar(CharPath & UserName & ".chr", "FACTION", "EXFACTION", CStr(Status))
142                         Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                        Else
144                         Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                        End If

                    Else
146                     Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChaosLegionKick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChaosLegionKick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRoyalArmyKick_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil)) <> 0 Or .flags.PrivEspecial Then
            
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")
                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")
                End If

114             tUser = NameIndex(UserName)
            
116             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "ECHÓ DE LA REAL A: " & UserName)
            
118             If tUser > 0 Then
            
120                 If .Faction.Status > 0 Then
122                     Call mFacciones.Faction_RemoveUser(tUser)
124                     Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
126                     Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
128                     Call FlushBuffer(tUser)
                    Else
130                     Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else
            
132                 If FileExist(CharPath & UserName & ".chr") Then
                
                        Dim Status As Byte

134                     Status = val(GetVar(CharPath & UserName & ".chr", "FACTION", "STATUS"))
                    
136                     If Status > 0 Then
138                         Call WriteVar(CharPath & UserName & ".chr", "FACTION", "STATUS", 0)
140                         Call WriteVar(CharPath & UserName & ".chr", "FACTION", "EXFACTION", CStr(Status))
142                         Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                        Else
144                         Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                        End If

                    Else
146                     Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If

                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRoyalArmyKick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRoyalArmyKick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleForceMIDIAll_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
    
100     With UserList(UserIndex)

            Dim midiID As Byte

102         midiID = Reader.ReadInt()
        
104         If Not EsGmPriv(UserIndex) Then Exit Sub
        
106         'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast música: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
108         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMusic(midiID))
        End With

        '<EhFooter>
        Exit Sub

HandleForceMIDIAll_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleForceMIDIAll " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleForceWAVEAll_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
    
100     With UserList(UserIndex)

            Dim waveID As Byte

102         waveID = Reader.ReadInt()
        
104         If Not EsGmPriv(UserIndex) Then Exit Sub
        
106         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(waveID, NO_3D_SOUND, NO_3D_SOUND))
        End With

        '<EhFooter>
        Exit Sub

HandleForceWAVEAll_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleForceWAVEAll " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTileBlockedToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
            If Not EsGm(UserIndex) Then Exit Sub
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BLOQ")
        
106         If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
108             MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
            Else
110             MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
            End If
        
112         Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
        End With

        '<EhFooter>
        Exit Sub

HandleTileBlockedToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTileBlockedToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleKillNPCNoRespawn_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGm(UserIndex) Then Exit Sub
        
104         If .flags.TargetNPC = 0 Then Exit Sub
        
106         If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Pretoriano Then Exit Sub
        
108         If isNPCResucitador(.flags.TargetNPC) Then
110             Call DeleteAreaResuTheNpc(.flags.TargetNPC)
            End If
        
112         Call QuitarNPC(.flags.TargetNPC)
114         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/MATA " & Npclist(.flags.TargetNPC).Name)
        End With

        '<EhFooter>
        Exit Sub

HandleKillNPCNoRespawn_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleKillNPCNoRespawn " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleKillAllNearbyNPCs_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmDios(UserIndex) Then Exit Sub
        
            Dim X As Long

            Dim Y As Long
        
104         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
106             For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

108                 If X > 0 And Y > 0 And X < 101 And Y < 101 Then
110                     If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                    End If

112             Next X
114         Next Y

116         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/MASSKILL")
        End With

        '<EhFooter>
        Exit Sub

HandleKillAllNearbyNPCs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleKillAllNearbyNPCs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLastIP_Err
        '</EhHeader>

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName   As String

            Dim lista      As String

            Dim LoopC      As Byte

            Dim Priv       As Integer

            Dim validCheck As Boolean
        
102         Priv = PlayerType.Admin Or PlayerType.Dios
104         UserName = Reader.ReadString8()
        
106         If (.flags.Privilegios And (PlayerType.Admin)) <> 0 Then

                'Handle special chars
108             If (InStrB(UserName, "\") <> 0) Then
110                 UserName = Replace(UserName, "\", "")
                End If

112             If (InStrB(UserName, "\") <> 0) Then
114                 UserName = Replace(UserName, "/", "")
                End If

116             If (InStrB(UserName, "+") <> 0) Then
118                 UserName = Replace(UserName, "+", " ")
                End If
            
                'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
120             If NameIndex(UserName) > 0 Then
122                 validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
                Else
124                 validCheck = (UserDarPrivilegioLevel(UserName) And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
                End If
            
126             If validCheck Then
128                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/LASTIP " & UserName)
                
130                 If FileExist(CharPath & UserName & ".chr", vbNormal) Then
132                     lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"

134                     For LoopC = 1 To 5
136                         lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
138                     Next LoopC

140                     Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                    Else
142                     Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else
144                 Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleLastIP_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLastIP " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChatColor_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the user`s chat color
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Color As Long
        
102         Color = RGB(Reader.ReadInt(), Reader.ReadInt(), Reader.ReadInt())
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
106             .flags.ChatColor = Color
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleChatColor_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChatColor " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleIgnored_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Ignore the user
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
104             .flags.AdminPerseguible = Not .flags.AdminPerseguible
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleIgnored_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleIgnored " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSaveChars_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Save the characters
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmDios(UserIndex) Then Exit Sub
        
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha guardado todos los chars.")
        
106         Call DistributeExpAndGldGroups
108         Call GuardarUsuarios(False)
        End With

        '<EhFooter>
        Exit Sub

HandleSaveChars_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSaveChars " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoBackup_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the backup`s info of the map
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim doTheBackUp As Boolean
        
102         doTheBackUp = Reader.ReadBool()
        
104         If (.flags.Privilegios And (PlayerType.Admin)) = 0 Then Exit Sub
        
106         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre el BackUp.")
        
            'Change the boolean to byte in a fast way
108         If doTheBackUp Then
110             MapInfo(.Pos.Map).BackUp = 1
            Else
112             MapInfo(.Pos.Map).BackUp = 0
            End If
        
            'Change the boolean to string in a fast way
114         Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
116         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoBackup_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoBackup " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoPK_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the pk`s info of the  map
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim isMapPk As Boolean
        
102         isMapPk = Reader.ReadBool()
        
104         If Not EsGmDios(UserIndex) Then Exit Sub
        
106         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si es PK el mapa.")
        
108         MapInfo(.Pos.Map).Pk = isMapPk
        
            'Change the boolean to string in a fast way
110         Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

112         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoPK_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoPK " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleChangeMapInfoLvl(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoLvl_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        'Restringido de Nivel -> Options: Todo nivel disponible.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Elv As Byte
        
102         Elv = Reader.ReadInt()

104         If EsGmDios(UserIndex) Then
106             If (Elv > STAT_MAXELV + 1) Then
108                 Call WriteConsoleMsg(UserIndex, "El nivel máximo que puedes elegir es el máximo del juego +1", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
110             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si es restringido por nivel el mapa.")
                
112             MapInfo(UserList(UserIndex).Pos.Map).LvlMin = Elv
                
114             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "LvlMin", Elv)
116             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RestringidoNivel: " & MapInfo(.Pos.Map).LvlMin, FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoLvl_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoLvl " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleChangeMapInfoLimpieza(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoLimpieza_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        'Restringido de Limpieza -> Options: Si/No
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Value As Byte
        
102         Value = Reader.ReadInt()

104         If EsGmPriv(UserIndex) Then
            
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si afecta la limpieza en el mapa o no")
                
108             MapInfo(UserList(UserIndex).Pos.Map).Limpieza = Value
                
110             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Limpieza", Value)
112             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Limpieza: " & IIf((MapInfo(.Pos.Map).Limpieza = 1), "SI", "NO"), FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoLimpieza_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoLimpieza " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleChangeMapInfoItems(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoItems_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        'Restringido de Items -> Options: Si/No
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Value As Byte
        
102         Value = Reader.ReadInt()

104         If EsGmDios(UserIndex) Then
            
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si caen items o no")
                
108             MapInfo(UserList(UserIndex).Pos.Map).CaenItems = Value
                
110             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "CaenItems", Value)
112             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Caen Items: " & IIf((MapInfo(.Pos.Map).CaenItems = 1), "SI", "NO"), FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleChangeMapInfoExp(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleChangeMapInfoExp_Err

        '</EhHeader>

        Dim Exp As Single
    
100     Exp = Reader.ReadReal32
    
102     With UserList(UserIndex)

104         If EsGmPriv(UserIndex) Then
                If Exp = 255 Then
                    Call CheckHappyHour
                    'frmMain.chkHappy.Value = IIf(HappyHour = True, 1, 0)
                ElseIf Exp = 254 Then
                    Call CheckPartyTime
                    'frmMain.chkParty.Value = IIf(PartyTime = True, 1, 0)
                Else
                
106                 MapInfo(.Pos.Map).Exp = Exp
                
108                 If Exp > 0 Then
110                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Experiencia aumentada en " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")" & " x" & CStr(Exp), FontTypeNames.FONTTYPE_USERPREMIUM))
                    Else
112                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La experiencia de " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")" & " ha vuelto a la normalidad.", FontTypeNames.FONTTYPE_USERPREMIUM))
                    End If

                End If
                
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Cambio de EXP x" & Exp & IIf(Exp <> 255, " en el mapa " & MapInfo(.Pos.Map).Name & "(" & .Pos.Map & ")", vbNullString))
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoExp_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoExp " & "at line " & Erl

        

        '</EhFooter>
End Sub

Private Sub HandleChangeMapInfoAttack(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoAttack_Err
        '</EhHeader>

        Dim Activado As Byte
    
100     Activado = Reader.ReadInt8
    
102     With UserList(UserIndex)
    
104         If EsGmDios(UserIndex) Then
106             MapInfo(.Pos.Map).FreeAttack = IIf((Activado = 0), False, True)
            
108             If Activado > 0 Then
110                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Libre ataque en " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")", FontTypeNames.FONTTYPE_USERPREMIUM))
                Else
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ataque limitado entre facciones en " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")", FontTypeNames.FONTTYPE_USERPREMIUM))
                End If
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoAttack_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoAttack " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoRestricted_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
        '***************************************************

        Dim tStr As String
    
100     With UserList(UserIndex)
        
102         tStr = Reader.ReadString8()
        
104         If EsGmPriv(UserIndex) Then
106             If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
108                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si es restringido el mapa.")
                
110                 MapInfo(UserList(UserIndex).Pos.Map).Restringir = RestrictStringToByte(tStr)
                
112                 Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeMapInfoRestricted_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoRestricted " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoNoMagic_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'MagiaSinEfecto -> Options: "1" , "0".
        '***************************************************
    
        Dim nomagic As Boolean
    
100     With UserList(UserIndex)
        
102         nomagic = Reader.ReadBool
        
104         If EsGmDios(UserIndex) Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
108             MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
110             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
112             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoNoMagic_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoNoMagic " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoNoInvi_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'InviSinEfecto -> Options: "1", "0"
        '***************************************************
    
        Dim noinvi As Boolean
    
100     With UserList(UserIndex)
        
102         noinvi = Reader.ReadBool()
        
104         If EsGmDios(UserIndex) Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
108             MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
110             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
112             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoNoInvi_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoNoInvi " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoNoResu_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'ResuSinEfecto -> Options: "1", "0"
        '***************************************************
    
        Dim noresu As Boolean
    
100     With UserList(UserIndex)
        
102         noresu = Reader.ReadBool()
        
104         If EsGmDios(UserIndex) <> 0 Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
108             MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
110             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
112             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoNoResu_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoNoResu " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoLand_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
        '***************************************************

        Dim tStr As String
    
100     With UserList(UserIndex)
        
102         tStr = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin)) <> 0 Then
106             If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
108                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información del terreno del mapa.")
                
110                 MapInfo(UserList(UserIndex).Pos.Map).Terreno = TerrainStringToByte(tStr)
                
112                 Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.Map).Terreno), FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
118                 Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeMapInfoLand_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoLand " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoZone_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
        '***************************************************

        Dim tStr As String
    
100     With UserList(UserIndex)
        
102         tStr = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin)) <> 0 Then
106             If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
108                 Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información de la zona del mapa.")
110                 MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
112                 Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
118                 Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeMapInfoZone_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoZone " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
            
''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoStealNpc_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/07/2010
        'RoboNpcsPermitido -> Options: "1", "0"
        '***************************************************
    
        Dim RoboNpc As Byte
    
100     With UserList(UserIndex)
        
102         RoboNpc = val(IIf(Reader.ReadBool(), 1, 0))
        
104         If EsGmDios(UserIndex) Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido robar npcs en el mapa.")
            
108             MapInfo(UserList(UserIndex).Pos.Map).RoboNpcsPermitido = RoboNpc
            
110             Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
112             Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapInfo(.Pos.Map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleChangeMapInfoStealNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoStealNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
            
''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoNoOcultar_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 18/09/2010
        'OcultarSinEfecto -> Options: "1", "0"
        '***************************************************
    
        Dim NoOcultar As Byte

        Dim mapa      As Integer
    
100     With UserList(UserIndex)
        
102         NoOcultar = val(IIf(Reader.ReadBool(), 1, 0))
        
104         If EsGmDios(UserIndex) Then
            
106             mapa = .Pos.Map
            
108             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido ocultarse en el mapa " & mapa & ".")
            
110             MapInfo(mapa).OcultarSinEfecto = NoOcultar
            
112             Call WriteVar(App.Path & MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "OcultarSinEfecto", NoOcultar)
114             Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeMapInfoNoOcultar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoNoOcultar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
           
''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeMapInfoNoInvocar_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 18/09/2010
        'InvocarSinEfecto -> Options: "1", "0"
        '***************************************************
    
        Dim NoInvocar As Byte

        Dim mapa      As Integer
    
100     With UserList(UserIndex)
        
102         NoInvocar = val(IIf(Reader.ReadBool(), 1, 0))
        
104         If EsGmDios(UserIndex) Then
            
106             mapa = .Pos.Map
            
108             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido invocar en el mapa " & mapa & ".")
            
110             MapInfo(mapa).InvocarSinEfecto = NoInvocar
            
112             Call WriteVar(App.Path & MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "InvocarSinEfecto", NoInvocar)
114             Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeMapInfoNoInvocar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeMapInfoNoInvocar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSaveMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Saves the map
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmDios(UserIndex) Then Exit Sub
        
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
106         Call GrabarMapa(.Pos.Map, Maps_FilePath & "WORLDBACKUP\Mapa" & .Pos.Map)
        
108         Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
        End With

        '<EhFooter>
        Exit Sub

HandleSaveMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSaveMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDoBackUp_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Show dobackup messages
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmPriv(UserIndex) Then Exit Sub
        
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha hecho un backup.")
        
106         Call ES.DoBackUp 'Sino lo confunde con la id del paquete
        End With

        '<EhFooter>
        Exit Sub

HandleDoBackUp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDoBackUp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCreateNPC_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 26/09/2010
        '26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
        '***************************************************
100     With UserList(UserIndex)
        
            Dim NpcIndex As Integer
        
102         NpcIndex = Reader.ReadInt()
        
104         If Not EsGmDios(UserIndex) Then Exit Sub
                
                
               If GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NAME") = vbNullString Then Exit Sub
               
               
106         If val(GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NPCTYPE")) = eNPCType.Pretoriano Or val(GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NPCTYPE")) = eNPCType.eCommerceChar Then
108             Call WriteConsoleMsg(UserIndex, "No puedes sumonear esta criatura. Revisa el numero de la misma. Gracias atentamente lautaro.", FontTypeNames.FONTTYPE_WARNING)
                  Exit Sub
            End If
        
110         NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
112         If NpcIndex <> 0 Then
114             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Sumoneó a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleCreateNPC_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCreateNPC " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleCreateNPCWithRespawn_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 26/09/2010
        '26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim NpcIndex As Integer
        
102         NpcIndex = Reader.ReadInt()
        
104         If NpcIndex > NumNpcs Then Exit Sub
        
            If Not EsGmPriv(UserIndex) Then Exit Sub
         
106         If val(GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NPCTYPE")) = eNPCType.Pretoriano Then
108             Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros que funcionan como guardines/pretorianos.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
        
110
        
112         NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
114         If NpcIndex <> 0 Then
116             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Sumoneó con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

            End If

        End With

        '<EhFooter>
        Exit Sub

HandleCreateNPCWithRespawn_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCreateNPCWithRespawn " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleServerOpenToUsersToggle_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If Not EsGmPriv(UserIndex) Then Exit Sub
        
104         If ServerSoloGMs > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
108             ServerSoloGMs = 0
110             frmServidor.chkServerHabilitado.Value = vbUnchecked
            Else

                Dim A As Long
                
                For A = 1 To LastUser

                    If Not EsGm(A) Then
                        Call Protocol.Kick(UserIndex, "Servidor restringido para administradores")

                    End If

                Next A
                
112             Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
114             ServerSoloGMs = 1
116             frmServidor.chkServerHabilitado.Value = vbChecked

            End If

        End With

        '<EhFooter>
        Exit Sub

HandleServerOpenToUsersToggle_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleServerOpenToUsersToggle " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTurnOffServer_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        'Turns off the server
        '***************************************************
        Dim handle As Integer
    
100     With UserList(UserIndex)
        
102         If Not EsGmPriv(UserIndex) Then Exit Sub
        
104         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/APAGAR")
106         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡¡" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
            'Log
108         handle = FreeFile
110         Open LogPath & "Main.log" For Append Shared As #handle
        
112         Print #handle, Date & " " & Time & " server apagado por " & .Name & ". "
        
114         Close #handle
        
116         Unload frmMain
        End With

        '<EhFooter>
        Exit Sub

HandleTurnOffServer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTurnOffServer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleTurnCriminal_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/26/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/CONDEN " & UserName)
            
108             tUser = NameIndex(UserName)

110             If tUser > 0 Then Call VolverCriminal(tUser)
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleTurnCriminal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleTurnCriminal " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleResetFactions_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 06/09/09
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim Char     As String

            Dim Temp     As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/RAJAR " & UserName)
            
108             tUser = NameIndex(UserName)
            
110             If tUser > 0 Then
112                 If UserList(tUser).Faction.Status = 0 Then
114                     Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                    Else
116                     Call mFacciones.Faction_RemoveUser(tUser)

                    End If

                Else
118                 Char = CharPath & UserName & ".chr"
                
120                 If FileExist(Char, vbNormal) Then
122                     Temp = val(GetVar(Char, "FACTION", "STATUS"))
                    
124                     If Temp = 0 Then
126                         Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                        Else
128                         Call WriteVar(Char, "FACTION", "STATUS", "0")
130                         Call WriteVar(Char, "FACTION", "ExFaction", CStr(Temp))
132                         Call WriteVar(Char, "FACTION", "StartDate", vbNullString)
134                         Call WriteVar(Char, "FACTION", "StartElv", "0")
136                         Call WriteVar(Char, "FACTION", "StartFrags", "0")
                                
                            Dim A As Long
                                
                            For A = 1 To MAX_INVENTORY_SLOTS

                                If .Invent.Object(A).ObjIndex > 0 Then
                                    If ObjData(.Invent.Object(A).ObjIndex).Real = 1 Or ObjData(.Invent.Object(A).ObjIndex).Caos = 1 Then
                                        Call QuitarObjetos(.Invent.Object(A).ObjIndex, .Invent.Object(A).Amount, UserIndex)

                                    End If

                                End If

                            Next A

                        End If
                    
                    Else
138                     Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleResetFactions_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleResetFactions " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSystemMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/29/06
        'Send a message to all the users
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
106             Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje de sistema:" & Message)
            
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleSystemMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSystemMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePing_Err
        '</EhHeader>

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Show ping messages
        '***************************************************

100     Call WritePong(UserIndex, Reader.ReadReal64())

        '<EhFooter>
        Exit Sub

HandlePing_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePing " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCreatePretorianClan_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        '***************************************************

        Dim Map   As Integer

        Dim X     As Byte

        Dim Y     As Byte

        Dim Index As Long
    
100     With UserList(UserIndex)
        
102         Map = Reader.ReadInt()
104         X = Reader.ReadInt()
106         Y = Reader.ReadInt()
        
108         If Not EsGmPriv(UserIndex) Then Exit Sub
        
            ' Valid pos?
110         If Not InMapBounds(Map, X, Y) Then
112             Call WriteConsoleMsg(UserIndex, "Posición inválida.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            ' Is already active any clan?
114         If Not ClanPretoriano(7).Active Then
            
116             If Not ClanPretoriano(Index).SpawnClan(Map, X, Y, Index) Then
118                 Call WriteConsoleMsg(UserIndex, "La posición no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)
                End If
        
            Else
120             Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)
            End If
    
        End With

        Exit Sub

122     Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.number & " - " & Err.description)
        '<EhFooter>
        Exit Sub

HandleCreatePretorianClan_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCreatePretorianClan " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDeletePretorianClan_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        '***************************************************
    
        Dim Map   As Integer

        Dim Index As Long
    
100     With UserList(UserIndex)
        
102         Map = Reader.ReadInt()
        
            ' User Admin?
104         If Not EsGmPriv(UserIndex) Then Exit Sub
        
            ' Valid map?
106         If Map < 1 Or Map > NumMaps Then
108             Call WriteConsoleMsg(UserIndex, "Mapa inválido.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            ' Search for the clan to be deleted
110         If ClanPretoriano(7).ClanMap = Map Then
112             ClanPretoriano(7).DeleteClan

            End If
    
        End With

        Exit Sub

114     Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.number & " - " & Err.description)
        '<EhFooter>
        Exit Sub

HandleDeletePretorianClan_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDeletePretorianClan " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "Logged" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteLoggedMessage_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Logged" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.logged)
    
102     With UserList(UserIndex)
    
104         Call Writer.WriteInt8(.Clase)
106         Call Writer.WriteInt8(.Raza)
            Call Writer.WriteInt8(.Genero)
108         Call Writer.WriteInt8(.Account.CharsAmount)
              Call Writer.WriteInt32(.Account.Gld)
        End With
    
110     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteLoggedMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteLoggedMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRemoveAllDialogs_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.RemoveDialogs)

102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteRemoveAllDialogs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteRemoveAllDialogs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal charindex As Integer)
        '<EhHeader>
        On Error GoTo WriteRemoveCharDialog_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageRemoveCharDialog(charindex))
        '<EhFooter>
        Exit Sub

WriteRemoveCharDialog_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteRemoveCharDialog " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteNavigateToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "NavigateToggle" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.NavigateToggle)

102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteNavigateToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteNavigateToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer, _
                           Optional ByVal Account As Boolean = False)
        '<EhHeader>
        On Error GoTo WriteDisconnect_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Disconnect" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.Disconnect)
102     Call Writer.WriteBool(Account)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteDisconnect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteDisconnect " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserOfferConfirm_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UserOfferConfirm)

102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUserOfferConfirm_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserOfferConfirm " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCommerceEnd_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CommerceEnd" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.CommerceEnd)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteCommerceEnd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCommerceEnd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBankEnd_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankEnd" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.BankEnd)

102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteBankEnd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteBankEnd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer, _
                             ByVal NpcName As String, _
                             ByVal Quest As Byte, _
                             ByRef QuestList() As Byte)
        '<EhHeader>
        On Error GoTo WriteCommerceInit_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CommerceInit" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.CommerceInit)
102     Call Writer.WriteString8(NpcName)
104     Call Writer.WriteInt8(Quest)
    
106     If Quest > 0 Then
108         Call Writer.WriteSafeArrayInt8(QuestList)
        End If
    
110     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteCommerceInit_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCommerceInit " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer, ByVal TypeBank As Byte)
        '<EhHeader>
        On Error GoTo WriteBankInit_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankInit" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.BankInit)
102     Call Writer.WriteInt(UserList(UserIndex).Account.Gld)
104     Call Writer.WriteInt(UserList(UserIndex).Account.Eldhir)
106     Call Writer.WriteInt(TypeBank)
    
108     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteBankInit_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteBankInit " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserCommerceInit_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UserCommerceInit)
102     Call Writer.WriteString8(UserList(UserIndex).ComUsu.DestNick)

104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUserCommerceInit_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserCommerceInit " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserCommerceEnd_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UserCommerceEnd)

102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUserCommerceEnd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserCommerceEnd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateSta_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateMana" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateSta)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.MinSta)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateSta_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateSta " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateMana_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateMana" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateMana)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.MinMan)
104     Call Writer.WriteInt16(1)
106     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateMana_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateMana " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateHP_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateMana" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateHP)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.MinHp)
104     Call Writer.WriteInt16(1)
106     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateHP_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateHP " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateGold_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateGold" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateGold)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.Gld)
        
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUpdateGold_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateGold " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateDsp" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDsp(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateDsp_Err
        '</EhHeader>

        '***************************************************
        'Author: WAICON
        'Last Modification: 06/05/2019
        'Writes the "UpdateDsp" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateDsp)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Eldhir)
        Call Writer.WriteInt32(UserList(UserIndex).Account.Eldhir)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateDsp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateDsp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateBankGold_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.UpdateBankGold)
102     Call Writer.WriteInt(UserList(UserIndex).Account.Gld)
104     Call Writer.WriteInt(UserList(UserIndex).Account.Eldhir)
106     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateBankGold_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateBankGold " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateExp_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateExp" message to the given user's outgoing data buffer
        '**************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateExp)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateExp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateExp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateStrenghtAndDexterity_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateStrenghtAndDexterity)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
104     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
106     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateStrenghtAndDexterity_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateStrenghtAndDexterity " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateDexterity_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateDexterity)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateDexterity_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateDexterity " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateStrenght_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateStrenght)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateStrenght_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateStrenght " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
        '<EhHeader>
        On Error GoTo WriteChangeMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeMap" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ChangeMap)
102     Call Writer.WriteInt(Map)
        
104     If Map <> 0 Then
106         Call Writer.WriteString8(MapInfo(Map).Name)
        Else
108         Call Writer.WriteString8(vbNullString)
        End If

110     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteChangeMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WritePosUpdate_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PosUpdate" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.PosUpdate)
102     Call Writer.WriteInt(UserList(UserIndex).Pos.X)
104     Call Writer.WriteInt(UserList(UserIndex).Pos.Y)
106     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WritePosUpdate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WritePosUpdate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal charindex As Integer, _
                             ByVal Color As Long)
        '<EhHeader>
        On Error GoTo WriteChatOverHead_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChatOverHead" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageChatOverHead(chat, charindex, Color))

        '<EhFooter>
        Exit Sub

WriteChatOverHead_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChatOverHead " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteChatPersonalizado(ByVal UserIndex As Integer, _
                                  ByVal chat As String, _
                                  ByVal charindex As Integer, _
                                  ByVal Tipo As Byte)
        '<EhHeader>
        On Error GoTo WriteChatPersonalizado_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Dalmasso (CHOTS)
        'Last Modification: 11/06/2011
        'Writes the "ChatPersonalizado" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageChatPersonalizado(chat, charindex, Tipo))

        '<EhFooter>
        Exit Sub

WriteChatPersonalizado_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChatPersonalizado " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
                           ByVal chat As String, _
                           ByVal FontIndex As FontTypeNames, _
                           Optional ByVal MessageType As eMessageType = Info)
        '<EhHeader>
        On Error GoTo WriteConsoleMsg_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageConsoleMsg(chat, FontIndex, MessageType))
    
        '<EhFooter>
        Exit Sub

WriteConsoleMsg_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteConsoleMsg " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal FontIndex As FontTypeNames)
        '<EhHeader>
        On Error GoTo WriteCommerceChat_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/17/06
        'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call SendData(ToOne, UserIndex, PrepareCommerceConsoleMsg(chat, FontIndex))
    
        '<EhFooter>
        Exit Sub

WriteCommerceChat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCommerceChat " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteShowMessageBox_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
102     Call Writer.WriteString8(Message)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteShowMessageBox_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteShowMessageBox " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserIndexInServer_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UserIndexInServer)
102     Call Writer.WriteInt(UserIndex)
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUserIndexInServer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserIndexInServer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserCharIndexInServer_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UserCharIndexInServer)
102     Call Writer.WriteInt(UserList(UserIndex).Char.charindex)
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUserCharIndexInServer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserCharIndexInServer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
                                ByVal Body As Integer, _
                                ByVal BodyAttack As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal charindex As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte, _
                                ByVal Weapon As Integer, _
                                ByVal Shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal Name As String, _
                                ByVal NickColor As Byte, _
                                ByVal Privileges As Byte, _
                                ByRef AuraIndex() As Byte, _
                                ByVal speeding As Single, _
                                ByVal Idle As Boolean, _
                                Optional ByVal NpcIndex As Integer = 0)
        '<EhHeader>
        On Error GoTo WriteCharacterCreate_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CharacterCreate" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call SendData(ToOne, UserIndex, PrepareMessageCharacterCreate(Body, BodyAttack, Head, Heading, charindex, X, Y, Weapon, Shield, FX, FXLoops, helmet, _
                                Name, NickColor, Privileges, AuraIndex, NpcIndex, Idle, False, speeding))

        '<EhFooter>
        Exit Sub

WriteCharacterCreate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCharacterCreate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal charindex As Integer)
        '<EhHeader>
        On Error GoTo WriteCharacterRemove_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CharacterRemove" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call SendData(ToOne, UserIndex, PrepareMessageCharacterRemove(charindex))

        '<EhFooter>
        Exit Sub

WriteCharacterRemove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCharacterRemove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal charindex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteCharacterMove_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CharacterMove" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageCharacterMove(charindex, X, Y))

        '<EhFooter>
        Exit Sub

WriteCharacterMove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCharacterMove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
        '<EhHeader>
        On Error GoTo WriteForceCharMove_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 26/03/2009
        'Writes the "ForceCharMove" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageForceCharMove(Direccion))

        '<EhFooter>
        Exit Sub

WriteForceCharMove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteForceCharMove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
                                ByVal Body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal charindex As Integer, _
                                ByVal Weapon As Integer, _
                                ByVal Shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByRef AuraIndex() As Byte, _
                                ByVal Idle As Boolean, _
                                ByVal Navegacion As Boolean)
        '<EhHeader>
        On Error GoTo WriteCharacterChange_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CharacterChange" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageCharacterChange(Body, 0, Head, Heading, charindex, Weapon, Shield, FX, FXLoops, helmet, AuraIndex, UserList(UserIndex).flags.ModoStream, Idle, Navegacion))

        '<EhFooter>
        Exit Sub

WriteCharacterChange_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCharacterChange " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, _
                             ByVal ObjIndex As Integer, _
                             ByVal GrhIndex As Long, _
                             ByVal X As Byte, _
                             ByVal Y As Byte, _
                             ByVal Sound As Integer)
        '<EhHeader>
        On Error GoTo WriteObjectCreate_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ObjectCreate" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageObjectCreate(ObjIndex, GrhIndex, X, Y, vbNullString, 0, Sound))

        '<EhFooter>
        Exit Sub

WriteObjectCreate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteObjectCreate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteObjectDelete_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ObjectDelete" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageObjectDelete(X, Y))

        '<EhFooter>
        Exit Sub

WriteObjectDelete_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteObjectDelete " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Boolean)
        '<EhHeader>
        On Error GoTo WriteBlockPosition_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BlockPosition" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.BlockPosition)
102     Call Writer.WriteInt(X)
104     Call Writer.WriteInt(Y)
106     Call Writer.WriteBool(Blocked)

108     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteBlockPosition_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteBlockPosition " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "PlayMusic" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMusic(ByVal UserIndex As Integer, ByVal Music As Integer)
        '<EhHeader>
        On Error GoTo WritePlayMusic_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PlayMusic" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessagePlayMusic(Music))

        '<EhFooter>
        Exit Sub

WritePlayMusic_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WritePlayMusic " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "PlayEffect" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayEffect(ByVal UserIndex As Integer, _
                         ByVal Wave As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WritePlayEffect_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Wave, X, Y))

        '<EhFooter>
        Exit Sub

WritePlayEffect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WritePlayEffect " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WritePauseToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PauseToggle" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessagePauseToggle())

        '<EhFooter>
        Exit Sub

WritePauseToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WritePauseToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
                         ByVal charindex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)
        '<EhHeader>
        On Error GoTo WriteCreateFX_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreateFX" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageCreateFX(charindex, FX, FXLoops))

        '<EhFooter>
        Exit Sub

WriteCreateFX_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCreateFX " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateUserStats_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateUserStats)
102     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
106     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMan)
108     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMan)
110     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
112     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
114     Call Writer.WriteInt32(UserList(UserIndex).Stats.Gld)
116     Call Writer.WriteInt32(UserList(UserIndex).Stats.Eldhir)
118     Call Writer.WriteInt8(UserList(UserIndex).Stats.Elv)
120     Call Writer.WriteInt32(UserList(UserIndex).Stats.Elu)
122     Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
124     Call Writer.WriteInt32(UserList(UserIndex).Stats.Points)
    
        Dim estatus As Byte
    
126     If UserList(UserIndex).flags.Bronce = 1 Then
128         estatus = 1
        End If
    
130     If UserList(UserIndex).flags.Plata = 1 Then
132         estatus = 2
        End If
    
134     If UserList(UserIndex).flags.Oro = 1 Then
136         estatus = 3
        End If
    
138     Call Writer.WriteInt8(estatus)
140     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateUserStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateUserStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteChangeInventorySlot_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 25/05/2011 (Amraphen)
        'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
        '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
        '25/05/2011: Amraphen - Ahora se envía la defensa según se tiene equipado armadura de segunda jerarquía o no.
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ChangeInventorySlot)
102     Call Writer.WriteInt(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
104     ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
106     Call Writer.WriteInt(ObjIndex)
        
108     If ObjIndex > 0 Then
110         obData = ObjData(ObjIndex)
        
            'Si tiene armadura de segunda jerarquía obtiene un porcentaje de defensa adicional.
112         If obData.Caos = 1 Or obData.Real = 1 Then
114             If UserList(UserIndex).Faction.Status > 0 Then
116                 obData.MinDef = obData.MinDef + InfoFaction(UserList(UserIndex).Faction.Status).Range(UserList(UserIndex).Faction.Range).MinDef
118                 obData.MaxDef = obData.MaxDef + InfoFaction(UserList(UserIndex).Faction.Status).Range(UserList(UserIndex).Faction.Range).MaxDef
                End If
            End If
                
        End If
        
120     Call Writer.WriteString8(obData.Name)
122     Call Writer.WriteInt(UserList(UserIndex).Invent.Object(Slot).Amount)
124     Call Writer.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
126     Call Writer.WriteInt(obData.GrhIndex)
128     Call Writer.WriteInt(obData.OBJType)
130     Call Writer.WriteInt(obData.MaxHit)
132     Call Writer.WriteInt(obData.MinHit)
134     Call Writer.WriteInt(obData.MaxDef)
136     Call Writer.WriteInt(obData.MinDef)
138     Call Writer.WriteReal32(SalePrice(ObjIndex))
140     Call Writer.WriteReal32(SalePriceDiamanteAzul(ObjIndex))
142     Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
    
144     Call Writer.WriteInt(obData.MinHitMag)
146     Call Writer.WriteInt(obData.MaxHitMag)
148     Call Writer.WriteInt(obData.DefensaMagicaMin)
150     Call Writer.WriteInt(obData.DefensaMagicaMax)
    
152     Call Writer.WriteInt8(obData.Bronce)
154     Call Writer.WriteInt8(obData.Plata)
156     Call Writer.WriteInt8(obData.Oro)
158     Call Writer.WriteInt8(obData.Premium)
    
160     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteChangeInventorySlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeInventorySlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)
        '<EhHeader>
        On Error GoTo WriteAddSlots_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 01/12/09
        'Writes the "AddSlots" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.AddSlots)
102     Call Writer.WriteInt(Mochila)
    
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteAddSlots_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteAddSlots " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteChangeBankSlot_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/03/09
        'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
        '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ChangeBankSlot)
102     Call Writer.WriteInt(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
104     ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
106     Call Writer.WriteInt(ObjIndex)
        
108     If ObjIndex > 0 Then
110         obData = ObjData(ObjIndex)
        End If
        
112     Call Writer.WriteString8(obData.Name)
114     Call Writer.WriteInt(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
116     Call Writer.WriteInt(obData.GrhIndex)
118     Call Writer.WriteInt(obData.OBJType)
120     Call Writer.WriteInt(obData.MaxHit)
122     Call Writer.WriteInt(obData.MinHit)
124     Call Writer.WriteInt(obData.MaxDef)
126     Call Writer.WriteInt(obData.MinDef)
128     Call Writer.WriteInt(obData.Valor)
130     Call Writer.WriteInt(obData.ValorEldhir)
132     Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
        
134     Call Writer.WriteInt(obData.MinHitMag)
136     Call Writer.WriteInt(obData.MaxHitMag)
138     Call Writer.WriteInt(obData.DefensaMagicaMin)
140     Call Writer.WriteInt(obData.DefensaMagicaMax)

142     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteChangeBankSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeBankSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeBankSlot_Account" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot_Account(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteChangeBankSlot_Account_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/03/09
        'Writes the "ChangeBankSlot_Account" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ChangeBankSlot_Account)
102     Call Writer.WriteInt(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
104     ObjIndex = UserList(UserIndex).Account.BancoInvent.Object(Slot).ObjIndex
        
106     Call Writer.WriteInt(ObjIndex)
        
108     If ObjIndex > 0 Then
110         obData = ObjData(ObjIndex)
        End If
        
112     Call Writer.WriteString8(obData.Name)
114     Call Writer.WriteInt(UserList(UserIndex).Account.BancoInvent.Object(Slot).Amount)
116     Call Writer.WriteInt(obData.GrhIndex)
118     Call Writer.WriteInt(obData.OBJType)
120     Call Writer.WriteInt(obData.MaxHit)
122     Call Writer.WriteInt(obData.MinHit)
124     Call Writer.WriteInt(obData.MaxDef)
126     Call Writer.WriteInt(obData.MinDef)
128     Call Writer.WriteInt(obData.Valor)
130     Call Writer.WriteInt(obData.ValorEldhir)
132     Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
        
134     Call Writer.WriteInt(obData.MinHitMag)
136     Call Writer.WriteInt(obData.MaxHitMag)
138     Call Writer.WriteInt(obData.DefensaMagicaMin)
140     Call Writer.WriteInt(obData.DefensaMagicaMax)

142     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteChangeBankSlot_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeBankSlot_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
        '<EhHeader>
        On Error GoTo WriteChangeSpellSlot_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ChangeSpellSlot)
102     Call Writer.WriteInt(Slot)
104     Call Writer.WriteInt(UserList(UserIndex).Stats.UserHechizos(Slot))
        
106     If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
108         Call Writer.WriteString8(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
        Else
110         Call Writer.WriteString8("(Vacio)")
        End If

112     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteChangeSpellSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeSpellSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteAttributes_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Atributes" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.Atributes)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
104     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
106     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
108     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
110     Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    
112     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteAttributes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteAttributes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


''
' Writes the "RestOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRestOK_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RestOK" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.RestOK)

102     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteRestOK_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteRestOK " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteErrorMsg_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ErrorMsg" message to the given user's outgoing data buffer
        '***************************************************

100     Call SendData(ToOne, UserIndex, PrepareMessageErrorMsg(Message), , True)

        '<EhFooter>
        Exit Sub

WriteErrorMsg_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteErrorMsg " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "Blind" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBlind_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Blind" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.Blind)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteBlind_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteBlind " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDumb_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Dumb" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.Dumb)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteDumb_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteDumb " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
                                       ByVal Slot As Byte, _
                                       ByRef Obj As Obj, _
                                       ByVal Price As Single, _
                                       ByVal Price2 As Single)
        '<EhHeader>
        On Error GoTo WriteChangeNPCInventorySlot_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/03/09
        'Last Modified by: Budi
        'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
        '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
        '***************************************************

        Dim ObjInfo As ObjData
    
100     If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
102         ObjInfo = ObjData(Obj.ObjIndex)
        
        End If
    
104     Call Writer.WriteInt(ServerPacketID.ChangeNPCInventorySlot)
106     Call Writer.WriteInt(Slot)
108     Call Writer.WriteInt(Obj.ObjIndex)
110     Call Writer.WriteString8(ObjInfo.Name)
112     Call Writer.WriteInt(Obj.Amount)
114     Call Writer.WriteReal32(Price)
116     Call Writer.WriteInt(ObjInfo.GrhIndex)
            
118     Call Writer.WriteInt(ObjInfo.OBJType)
120     Call Writer.WriteInt(ObjInfo.MaxHit)
122     Call Writer.WriteInt(ObjInfo.MinHit)
124     Call Writer.WriteInt(ObjInfo.MaxDef)
126     Call Writer.WriteInt(ObjInfo.MinDef)
128     Call Writer.WriteReal32(Price2)
130     Call Writer.WriteBool(CanUse_Inventory(UserIndex, Obj.ObjIndex))
            
132     Call Writer.WriteInt(ObjInfo.MinHitMag)
134     Call Writer.WriteInt(ObjInfo.MaxHitMag)
136     Call Writer.WriteInt(ObjInfo.DefensaMagicaMin)
138     Call Writer.WriteInt(ObjInfo.DefensaMagicaMax)
        
140     Call Writer.WriteInt(NpcInventory_GetAnimation(UserIndex, Obj.ObjIndex))
    
142     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteChangeNPCInventorySlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeNPCInventorySlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateHungerAndThirst_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateHungerAndThirst)
102     Call Writer.WriteInt(UserList(UserIndex).Stats.MaxAGU)
104     Call Writer.WriteInt(UserList(UserIndex).Stats.MinAGU)
106     Call Writer.WriteInt(UserList(UserIndex).Stats.MaxHam)
108     Call Writer.WriteInt(UserList(UserIndex).Stats.MinHam)
        
110     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateHungerAndThirst_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateHungerAndThirst " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


''
' Writes the "MiniStats" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer, ByVal tUser As Integer)
        '<EhHeader>
        On Error GoTo WriteMiniStats_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "MiniStats" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.MiniStats)
        
102     Call Writer.WriteInt16(UserList(tUser).Faction.FragsCiu)
104     Call Writer.WriteInt16(UserList(tUser).Faction.FragsCri)
        
        
106     Call Writer.WriteInt8(UserList(tUser).Clase)
108     Call Writer.WriteInt8(UserList(tUser).Raza)
110     Call Writer.WriteInt32(UserList(tUser).Reputacion.promedio)
    
112     Call Writer.WriteInt8(UserList(tUser).Stats.Elv)
114     Call Writer.WriteInt32(UserList(tUser).Stats.Exp)
116     Call Writer.WriteInt32(UserList(tUser).Stats.Elu)
118     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteMiniStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMiniStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data Reader.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
        '<EhHeader>
        On Error GoTo WriteLevelUp_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "LevelUp" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.LevelUp)
102     Call Writer.WriteInt(skillPoints)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteLevelUp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteLevelUp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
                             ByVal charindex As Integer, _
                             ByVal Invisible As Boolean)
        '<EhHeader>
        On Error GoTo WriteSetInvisible_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SetInvisible" message to the given user's outgoing data buffer
        '***************************************************
    
100     Call SendData(ToOne, UserIndex, PrepareMessageSetInvisible(charindex, Invisible))

        '<EhFooter>
        Exit Sub

WriteSetInvisible_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteSetInvisible " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


''
' Writes the "BlindNoMore" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBlindNoMore_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BlindNoMore" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.BlindNoMore)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteBlindNoMore_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteBlindNoMore " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDumbNoMore_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "DumbNoMore" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.DumbNoMore)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteDumbNoMore_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteDumbNoMore " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSendSkills_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 11/19/09
        'Writes the "SendSkills" message to the given user's outgoing data buffer
        '11/19/09: Pato - Now send the percentage of progress of the skills.
        '***************************************************

        Dim i As Long
    
100     With UserList(UserIndex)
102         Call Writer.WriteInt(ServerPacketID.SendSkills)
        
104         Call Writer.WriteInt8(UserList(UserIndex).Clase)
        
106         For i = 1 To NUMSKILLS
108             Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(i))
110         Next i

        End With

112     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteSendSkills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteSendSkills " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteParalizeOK_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/12/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Writes the "ParalizeOK" message to the given user's outgoing data buffer
        'And updates user position
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ParalizeOK)
102     Call WritePosUpdate(UserIndex)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteParalizeOK_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteParalizeOK " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
        '<EhHeader>
        On Error GoTo WriteShowUserRequest_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ShowUserRequest)
102     Call Writer.WriteString8(details)
    
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteShowUserRequest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteShowUserRequest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteTradeOK_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TradeOK" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.TradeOK)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteTradeOK_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteTradeOK " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBankOK_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankOK" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.BankOK)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteBankOK_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteBankOK " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
                                    ByVal OfferSlot As Byte, _
                                    ByVal ObjIndex As Integer, _
                                    ByVal Amount As Long)
        '<EhHeader>
        On Error GoTo WriteChangeUserTradeSlot_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/03/09
        'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
        '25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
        '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ChangeUserTradeSlot)
        
102     Call Writer.WriteInt(OfferSlot)
104     Call Writer.WriteInt(ObjIndex)
106     Call Writer.WriteInt(Amount)
        
108     If ObjIndex > 0 Then
110         Call Writer.WriteInt(ObjData(ObjIndex).GrhIndex)
112         Call Writer.WriteInt(ObjData(ObjIndex).OBJType)
114         Call Writer.WriteInt(ObjData(ObjIndex).MaxHit)
116         Call Writer.WriteInt(ObjData(ObjIndex).MinHit)
118         Call Writer.WriteInt(ObjData(ObjIndex).MaxDef)
120         Call Writer.WriteInt(ObjData(ObjIndex).MinDef)
122         Call Writer.WriteInt(SalePrice(ObjIndex))
124         Call Writer.WriteString8(ObjData(ObjIndex).Name)
126         Call Writer.WriteInt(SalePriceDiamanteAzul(ObjIndex))
            
128         Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
130         Call Writer.WriteInt(ObjData(ObjIndex).Bronce)
132         Call Writer.WriteInt(ObjData(ObjIndex).Plata)
134         Call Writer.WriteInt(ObjData(ObjIndex).Oro)
136         Call Writer.WriteInt(ObjData(ObjIndex).Premium)
        End If

138     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteChangeUserTradeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteChangeUserTradeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
        '<EhHeader>
        On Error GoTo WriteSpawnList_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SpawnList" message to the given user's outgoing data buffer
        '***************************************************

        Dim i   As Long

        Dim Tmp As String
    
100     Call Writer.WriteInt(ServerPacketID.SpawnList)
        
102     For i = LBound(npcNames()) To UBound(npcNames())
104         Tmp = Tmp & npcNames(i) & SEPARATOR
106     Next i
        
108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
110     Call Writer.WriteString8(Tmp)

112     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteSpawnList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteSpawnList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



''
' Writes the "ShowDenounces" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenounces(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowDenounces_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        'Writes the "ShowDenounces" message to the given user's outgoing data buffer
        '***************************************************
    
        Dim DenounceIndex As Long

        Dim DenounceList  As String

100     Call Writer.WriteInt(ServerPacketID.ShowDenounces)
        
102     For DenounceIndex = 1 To Denuncias.Longitud
104         DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
106     Next DenounceIndex
        
108     If LenB(DenounceList) <> 0 Then DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
        
110     Call Writer.WriteString8(DenounceList)
        
112     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteShowDenounces_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteShowDenounces " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowGMPanelForm_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ShowGMPanelForm)
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteShowGMPanelForm_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteShowGMPanelForm " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
                             ByRef userNamesList() As String, _
                             ByVal cant As Integer)
        '<EhHeader>
        On Error GoTo WriteUserNameList_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06 NIGO:
        'Writes the "UserNameList" message to the given user's outgoing data buffer
        '***************************************************

        Dim i   As Long

        Dim Tmp As String
    
100     Call Writer.WriteInt(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
102     For i = 1 To cant
104         Tmp = Tmp & userNamesList(i) & SEPARATOR
106     Next i
        
108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
110     Call Writer.WriteString8(Tmp)
112     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUserNameList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserNameList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "Pong" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal tPing As Double)
        '<EhHeader>
        On Error GoTo WritePong_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Pong" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.Pong)
102     Call Writer.WriteReal64(tPing)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WritePong_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WritePong " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo FlushBuffer_Err
        '</EhHeader>
    
100     Server.Flush UserIndex

        '<EhFooter>
        Exit Sub

FlushBuffer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.FlushBuffer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal charindex As Integer, _
                                           ByVal Invisible As Boolean, _
                                           Optional ByVal Intermitencia As Boolean = False) As String
        '<EhHeader>
        On Error GoTo PrepareMessageSetInvisible_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "SetInvisible" message and returns it.
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.SetInvisible)
        
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteBool(Invisible)
106     Call Writer.WriteBool(Intermitencia)
        
        '<EhFooter>
        Exit Function

PrepareMessageSetInvisible_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageSetInvisible " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal charindex As Integer, _
                                                  ByVal NewNick As String) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterChangeNick_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 07/23/09
        'Prepares the "Change Nick" message and returns it.
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterChangeNick)
        
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteString8(NewNick)

        '<EhFooter>
        Exit Function

PrepareMessageCharacterChangeNick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterChangeNick " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, _
                                           ByVal charindex As Integer, _
                                           ByVal Color As Long) As String
        '<EhHeader>
        On Error GoTo PrepareMessageChatOverHead_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ChatOverHead" message and returns it.
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.ChatOverHead)
102     Call Writer.WriteString16(chat)
104     Call Writer.WriteInt(charindex)
        
        ' Write rgb channels and save one byte from long :D
106     Call Writer.WriteInt(Color And &HFF)
108     Call Writer.WriteInt((Color And &HFF00&) \ &H100&)
110     Call Writer.WriteInt((Color And &HFF0000) \ &H10000)

        '<EhFooter>
        Exit Function

PrepareMessageChatOverHead_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageChatOverHead " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PrepareMessageChatPersonalizado(ByVal chat As String, _
                                                ByVal charindex As Integer, _
                                                ByVal Tipo As Byte) As String
        '<EhHeader>
        On Error GoTo PrepareMessageChatPersonalizado_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Dalmasso (CHOTS)
        'Last Modification: 11/06/2011
        'Prepares the "ChatPersonalizado" message and returns it.
        '**************************************************
        
100     Call Writer.WriteInt(ServerPacketID.ChatPersonalizado)
102     Call Writer.WriteString16(chat)
104     Call Writer.WriteInt(charindex)
        
        ' Write the type of message
        '1=normal
        '2=clan
        '3=party
        '4=gritar
        '5=palabras magicas
        '6=susurrar
106     Call Writer.WriteInt(Tipo)

        '<EhFooter>
        Exit Function

PrepareMessageChatPersonalizado_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageChatPersonalizado " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @param    MessageType type of console message (General, Guild, Party)
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, _
                                         ByVal FontIndex As FontTypeNames, _
                                         Optional ByVal MessageType As eMessageType = Info) As String
        '<EhHeader>
        On Error GoTo PrepareMessageConsoleMsg_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/05/11 (D'Artagnan)
        'Prepares the "MessageType" message and returns it.
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.ConsoleMsg)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt(FontIndex)
106     Call Writer.WriteInt(MessageType)
        
        '<EhFooter>
        Exit Function

PrepareMessageConsoleMsg_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageConsoleMsg " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PrepareCommerceConsoleMsg(ByRef chat As String, _
                                          ByVal FontIndex As FontTypeNames) As String
        '<EhHeader>
        On Error GoTo PrepareCommerceConsoleMsg_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 03/12/2009
        'Prepares the "CommerceConsoleMsg" message and returns it.
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.CommerceChat)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt(FontIndex)

        '<EhFooter>
        Exit Function

PrepareCommerceConsoleMsg_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareCommerceConsoleMsg " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal charindex As Integer, _
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer, _
                                       Optional ByVal IsMeditation As Boolean = False) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCreateFX_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.CreateFX)
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(FX)
106     Call Writer.WriteInt(FXLoops)
108     Call Writer.WriteBool(IsMeditation)

        '<EhFooter>
        Exit Function

PrepareMessageCreateFX_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCreateFX " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "PlayEffect" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayEffect(ByVal Wave As Integer, _
                                         ByVal X As Byte, _
                                         ByVal Y As Byte, _
                                         Optional ByVal Entity As Long = 0, _
                                         Optional ByVal Repeat As Boolean = False, _
                                         Optional ByVal MapOnly As Boolean = False) As String
        '<EhHeader>
        On Error GoTo PrepareMessagePlayEffect_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.PlayWave)
102     Call Writer.WriteInt(Wave)
104     Call Writer.WriteInt(X)
106     Call Writer.WriteInt(Y)
108     Call Writer.WriteInt(Entity)
110     Call Writer.WriteBool(Repeat)
        Call Writer.WriteBool(MapOnly)

        '<EhFooter>
        Exit Function

PrepareMessagePlayEffect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessagePlayEffect " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function PrepareMessageStopWaveMap(ByVal X As Byte, _
                                                                        ByVal Y As Byte, _
                                                                        ByVal Inmediatily As Boolean) As String

        On Error GoTo PrepareMessagePlayEffect_Err
        
100     Call Writer.WriteInt(ServerPacketID.StopWaveMap)
104     Call Writer.WriteInt(X)
106     Call Writer.WriteInt(Y)
          Call Writer.WriteBool(Inmediatily)

        '<EhFooter>
        Exit Function

PrepareMessagePlayEffect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessagePlayEffect " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String
        '<EhHeader>
        On Error GoTo PrepareMessageShowMessageBox_Err
        '</EhHeader>

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        'Prepares the "ShowMessageBox" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
102     Call Writer.WriteString8(chat)

        '<EhFooter>
        Exit Function

PrepareMessageShowMessageBox_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageShowMessageBox " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "PlayMusic" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMusic(ByVal Music As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessagePlayMusic_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "PlayMusic" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.PlayMusic)
102     Call Writer.WriteInt(Music)

        '<EhFooter>
        Exit Function

PrepareMessagePlayMusic_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessagePlayMusic " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
        '<EhHeader>
        On Error GoTo PrepareMessagePauseToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "PauseToggle" message and returns it
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.PauseToggle)

        '<EhFooter>
        Exit Function

PrepareMessagePauseToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessagePauseToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
        '<EhHeader>
        On Error GoTo PrepareMessageObjectDelete_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ObjectDelete" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.ObjectDelete)
102     Call Writer.WriteInt(X)
104     Call Writer.WriteInt(Y)

        '<EhFooter>
        Exit Function

PrepareMessageObjectDelete_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageObjectDelete " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, _
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Boolean) As String
        '<EhHeader>
        On Error GoTo PrepareMessageBlockPosition_Err
        '</EhHeader>

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        'Prepares the "BlockPosition" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.BlockPosition)
102     Call Writer.WriteInt(X)
104     Call Writer.WriteInt(Y)
106     Call Writer.WriteBool(Blocked)
    
        '<EhFooter>
        Exit Function

PrepareMessageBlockPosition_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageBlockPosition " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal ObjIndex As Integer, _
                                           ByVal GrhIndex As Long, _
                                           ByVal X As Byte, _
                                           ByVal Y As Byte, _
                                           ByVal Name As String, _
                                           ByVal Amount As Integer, _
                                           ByVal Sound As Integer) As String

        '<EhHeader>
        On Error GoTo PrepareMessageObjectCreate_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'prepares the "ObjectCreate" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.ObjectCreate)
102     Call Writer.WriteInt(X)
104     Call Writer.WriteInt(Y)
106     Call Writer.WriteInt(GrhIndex)
108     Call Writer.WriteInt16(ObjIndex)
110     Call Writer.WriteInt(Amount)
112     Call Writer.WriteInt16(Sound)

        '<EhFooter>
        Exit Function

PrepareMessageObjectCreate_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageObjectCreate " & "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal charindex As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterRemove_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterRemove" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterRemove)
102     Call Writer.WriteInt(charindex)

        '<EhFooter>
        Exit Function

PrepareMessageCharacterRemove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterRemove " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal charindex As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessageRemoveCharDialog_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.RemoveCharDialog)
102     Call Writer.WriteInt(charindex)

        '<EhFooter>
        Exit Function

PrepareMessageRemoveCharDialog_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageRemoveCharDialog " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data Reader.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, _
                                              ByVal BodyAttack As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal charindex As Integer, _
                                              ByVal X As Byte, _
                                              ByVal Y As Byte, _
                                              ByVal Weapon As Integer, _
                                              ByVal Shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Name As String, _
                                              ByVal NickColor As Byte, _
                                              ByVal Privileges As Byte, _
                                              ByRef AuraIndex() As Byte, _
                                              ByVal NpcIndex As Integer, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean, _
                                              ByVal speeding As Single) As String

        '<EhHeader>
        On Error GoTo PrepareMessageCharacterCreate_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterCreate" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterCreate)
        
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(Body)
106     Call Writer.WriteInt(BodyAttack)
108     Call Writer.WriteInt(Head)
110     Call Writer.WriteInt(Heading)
112     Call Writer.WriteInt(X)
114     Call Writer.WriteInt(Y)
116     Call Writer.WriteInt(Weapon)
118     Call Writer.WriteInt(Shield)
120     Call Writer.WriteInt(helmet)
122     Call Writer.WriteInt(FX)
124     Call Writer.WriteInt(FXLoops)
126     Call Writer.WriteString8(Name)
128     Call Writer.WriteInt(NickColor)
130     Call Writer.WriteInt(Privileges)
          
        Dim A As Long
          
        For A = 1 To MAX_AURAS
132         Call Writer.WriteInt(AuraIndex(A))
        Next A

134     Call Writer.WriteInt16(NpcIndex)

        Dim flags As Byte

        flags = 0
        
        If Idle Then flags = flags Or &O1 ' 00000001
        If Navegando Then flags = flags Or &O2
        Call Writer.WriteInt8(flags)
        Call Writer.WriteReal32(speeding)
        '<EhFooter>
        Exit Function

PrepareMessageCharacterCreate_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterCreate " & "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, _
                                              ByVal BodyAttack As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal charindex As Integer, _
                                              ByVal Weapon As Integer, _
                                              ByVal Shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByRef AuraIndex() As Byte, _
                                              ByVal ModoStreamer As Boolean, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterChange_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterChange" message and returns it
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.CharacterChange)
        
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(Body)
106     Call Writer.WriteInt(BodyAttack)
108     Call Writer.WriteInt(Head)
110     Call Writer.WriteInt(Heading)
112     Call Writer.WriteInt(Weapon)
114     Call Writer.WriteInt(Shield)
116     Call Writer.WriteInt(helmet)
118     Call Writer.WriteInt(FX)
120     Call Writer.WriteInt(FXLoops)
          
          
          Dim A As Long
          For A = 1 To MAX_AURAS
122         Call Writer.WriteInt(AuraIndex(A))
          Next A
          
124     Call Writer.WriteBool(ModoStreamer)

        Dim flags As Byte
        flags = 0
        If Idle Then flags = flags Or &O1
        If Navegando Then flags = flags Or &O2
        Call Writer.WriteInt8(flags)
        '<EhFooter>
        Exit Function

PrepareMessageCharacterChange_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterChange " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "CharacterChangeHeading" message and returns it.
'

Public Function PrepareMessageCharacterChangeHeading(ByVal charindex As Integer, _
                                                     ByVal Heading As eHeading) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterChangeHeading_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        'Prepares the "CharacterChangeHeading" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterChangeHeading)
        
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(Heading)

        '<EhFooter>
        Exit Function

PrepareMessageCharacterChangeHeading_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterChangeHeading " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal charindex As Integer, _
                                            ByVal X As Byte, _
                                            ByVal Y As Byte) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterMove_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterMove" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterMove)
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(X)
106     Call Writer.WriteInt(Y)

        '<EhFooter>
        Exit Function

PrepareMessageCharacterMove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterMove " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
        '<EhHeader>
        On Error GoTo PrepareMessageForceCharMove_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 26/03/2009
        'Prepares the "ForceCharMove" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.ForceCharMove)
102     Call Writer.WriteInt(Direccion)

        '<EhFooter>
        Exit Function

PrepareMessageForceCharMove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageForceCharMove " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
                                                 ByVal NickColor As Byte, _
                                                 ByRef Tag As String) As String
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateTagAndStatus_Err
        '</EhHeader>

        '***************************************************
        'Author: Alejandro Salvo (Salvito)
        'Last Modification: 04/07/07
        'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        'Prepares the "UpdateTagAndStatus" message and returns it
        '15/01/2010: ZaMa - Now sends the nick color instead of the status.
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.UpdateTagAndStatus)
        
102     Call Writer.WriteInt(UserList(UserIndex).Char.charindex)
104     Call Writer.WriteInt(NickColor)
106     Call Writer.WriteString8(Tag)
        
        '<EhFooter>
        Exit Function

PrepareMessageUpdateTagAndStatus_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageUpdateTagAndStatus " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String
        '<EhHeader>
        On Error GoTo PrepareMessageErrorMsg_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ErrorMsg" message and returns it
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.ErrorMsg)
102     Call Writer.WriteString8(Message)

        '<EhFooter>
        Exit Function

PrepareMessageErrorMsg_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageErrorMsg " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Writes the "CancelOfferItem" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteCancelOfferItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/03/2010
        '
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.CancelOfferItem)
102     Call Writer.WriteInt(Slot)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteCancelOfferItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCancelOfferItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "SetDialog" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetDialog(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSetDialog_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 18/11/2010
        '20/11/2010: ZaMa - Arreglo privilegios.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim NewDialog As String

102         NewDialog = Reader.ReadString8
        
104         If .flags.TargetNPC > 0 Then

                ' Dsgm/Dsrm/Rm
106             If EsGmPriv(UserIndex) Then
                    'Replace the NPC's dialog.
108                 Npclist(.flags.TargetNPC).Desc = NewDialog
                End If
            End If

        End With
    
        '<EhFooter>
        Exit Sub

HandleSetDialog_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSetDialog " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Impersonate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImpersonate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleImpersonate_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        '
        '***************************************************
100     With UserList(UserIndex)
        
            ' Dsgm/Dsrm/Rm
102         If Not EsGmPriv(UserIndex) Then Exit Sub
        
            Dim NpcIndex As Integer

104         NpcIndex = .flags.TargetNPC
        
106         If NpcIndex = 0 Then Exit Sub
        
            ' Copy head, body and desc
108         Call ImitateNpc(UserIndex, NpcIndex)
        
            ' Teleports user to npc's coords
110         Call WarpUserChar(UserIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, False, True)
        
            ' Log gm
112         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
            ' Remove npc
114         Call QuitarNPC(NpcIndex)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleImpersonate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleImpersonate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleImitate_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        '
        '***************************************************
    
100     If Not EsGmPriv(UserIndex) Then Exit Sub
    
102     With UserList(UserIndex)
            Dim NpcIndex As Integer

104         NpcIndex = .flags.TargetNPC
        
106         If NpcIndex = 0 Then Exit Sub
        
            ' Copy head, body and desc
108         Call ImitateNpc(UserIndex, NpcIndex)
110         Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleImitate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleImitate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RecordAdd" message.
'
' @param UserIndex The index of the user sending the message
           
Public Sub HandleRecordAdd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRecordAdd_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        '
        '**************************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim Reason   As String
        
102         UserName = Reader.ReadString8
104         Reason = Reader.ReadString8
    
106         If Not (.flags.Privilegios And (PlayerType.User)) Then

                'Verificamos que exista el personaje
108             If Not FileExist(CharPath & UCase$(UserName) & ".chr") Then
110                 Call WriteShowMessageBox(UserIndex, "El personaje no existe")
                Else
                    'Agregamos el seguimiento
112                 Call AddRecord(UserIndex, UserName, Reason)
                
                    'Enviamos la nueva lista de personajes
114                 Call WriteRecordList(UserIndex)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRecordAdd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRecordAdd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RecordAddObs" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordAddObs(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRecordAddObs_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        '
        '**************************************************************

100     With UserList(UserIndex)
        
            Dim RecordIndex As Byte

            Dim Obs         As String
        
102         RecordIndex = Reader.ReadInt
104         Obs = Reader.ReadString8
        
106         If Not (.flags.Privilegios And (PlayerType.User)) Then
                'Agregamos la observación
108             Call AddObs(UserIndex, RecordIndex, Obs)
            
                'Actualizamos la información
110             Call WriteRecordDetails(UserIndex, RecordIndex)
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRecordAddObs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRecordAddObs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RecordRemove" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordRemove(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRecordRemove_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        '
        '***************************************************
        Dim RecordIndex As Integer

100     With UserList(UserIndex)
    
102         RecordIndex = Reader.ReadInt
        
104         If .flags.Privilegios And (PlayerType.User) Then Exit Sub
        
            'Sólo dioses pueden remover los seguimientos, los otros reciben una advertencia:
106         If (.flags.Privilegios And PlayerType.Dios) Then
108             Call RemoveRecord(RecordIndex)
110             Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
112             Call WriteRecordList(UserIndex)
            Else
114             Call WriteShowMessageBox(UserIndex, "Sólo los dioses pueden eliminar seguimientos.")
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleRecordRemove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRecordRemove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RecordListRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordListRequest(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRecordListRequest_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        '
        '***************************************************
100     With UserList(UserIndex)

102         If .flags.Privilegios And (PlayerType.User) Then Exit Sub

104         Call WriteRecordList(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleRecordListRequest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRecordListRequest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "RecordDetails" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetails(ByVal UserIndex As Integer, ByVal RecordIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRecordDetails_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordDetails" message to the given user's outgoing data buffer
        '***************************************************
        Dim i        As Long

        Dim tIndex   As Integer

        Dim tmpStr   As String

        Dim TempDate As Date

100     Call Writer.WriteInt(ServerPacketID.RecordDetails)
        
        'Creador y motivo
102     Call Writer.WriteString8(Records(RecordIndex).Creador)
104     Call Writer.WriteString8(Records(RecordIndex).Motivo)
        
106     tIndex = NameIndex(Records(RecordIndex).Usuario)
        
        'Status del pj (online?)
108     Call Writer.WriteBool(tIndex > 0)
        
        'Escribo la IP según el estado del personaje
110     If tIndex > 0 Then
            'La IP Actual
112         tmpStr = UserList(tIndex).IpAddress
        Else 'String nulo
114         tmpStr = vbNullString
        End If

116     Call Writer.WriteString8(tmpStr)
        
        'Escribo tiempo online según el estado del personaje
118     If tIndex > 0 Then
            'Tiempo logueado.
120         TempDate = Now - UserList(tIndex).LogOnTime
122         tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
        Else
            'Envío string nulo.
124         tmpStr = vbNullString
        End If

126     Call Writer.WriteString8(tmpStr)

        'Escribo observaciones:
128     tmpStr = vbNullString

130     If Records(RecordIndex).NumObs Then

132         For i = 1 To Records(RecordIndex).NumObs
134             tmpStr = tmpStr & Records(RecordIndex).Obs(i).Creador & "> " & Records(RecordIndex).Obs(i).Detalles & vbCrLf
136         Next i
            
138         tmpStr = Left$(tmpStr, Len(tmpStr) - 1)
        End If

140     Call Writer.WriteString8(tmpStr)
142     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteRecordDetails_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteRecordDetails " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "RecordList" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordList(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRecordList_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordList" message to the given user's outgoing data buffer
        '***************************************************
        Dim i As Long
    
100     Call Writer.WriteInt(ServerPacketID.RecordList)
        
102     Call Writer.WriteInt(NumRecords)

104     For i = 1 To NumRecords
106         Call Writer.WriteString8(Records(i).Usuario)
108     Next i

110     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteRecordList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteRecordList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Writes the "ShowMenu" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    MenuIndex: The menu index.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMenu(ByVal UserIndex As Integer, ByVal MenuIndex As Byte)
        '<EhHeader>
        On Error GoTo WriteShowMenu_Err
        '</EhHeader>

        '***************************************************
        'Author: ZaMa
        'Last Modification: 10/05/2011
        'Writes the "ShowMenu" message to the given user's outgoing data buffer
        '***************************************************
        Dim i As Long

100     Call Writer.WriteInt(ServerPacketID.ShowMenu)
        
102     Call Writer.WriteInt(MenuIndex)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteShowMenu_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteShowMenu " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "RecordDetailsRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordDetailsRequest(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRecordDetailsRequest_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 07/04/2011
        'Handles the "RecordListRequest" message
        '***************************************************
        Dim RecordIndex As Byte

100     With UserList(UserIndex)
        
102         RecordIndex = Reader.ReadInt
        
104         If .flags.Privilegios And (PlayerType.User) Then Exit Sub
        
106         Call WriteRecordDetails(UserIndex, RecordIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleRecordDetailsRequest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRecordDetailsRequest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleMoveItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Ignacio Mariano Tirabasso (Budi)
        'Last Modification: 01/01/2011
        '
        '***************************************************
    
100     With UserList(UserIndex)

            Dim originalSlot As Byte

            Dim newSlot      As Byte
        
            Dim Tipo         As Byte
        
            Dim TypeBank     As Byte
        
102         originalSlot = Reader.ReadInt
104         newSlot = Reader.ReadInt
106         Tipo = Reader.ReadInt
108         TypeBank = Reader.ReadInt
        
110         If Tipo = eMoveType.Inventory Then
112             Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
114         ElseIf Tipo = eMoveType.Bank Then
116             Call InvUsuario.MoveItem_Bank(UserIndex, originalSlot, newSlot, TypeBank)
            
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleMoveItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleMoveItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function PrepareMessageCharacterAttackMovement(ByVal charindex As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterAttackMovement_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 24/05/2011
        'Prepares the "CharacterAttackMovement" message and returns it.
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterAttackMovement)
102     Call Writer.WriteInt(charindex)

        '<EhFooter>
        Exit Function

PrepareMessageCharacterAttackMovement_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterAttackMovement " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function PrepareMessageCharacterAttackNpc(ByVal charindex As Integer, _
                                                 ByVal BodyAttack As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterAttackNpc_Err
        '</EhHeader>

        '***************************************************
        'Author: Lautarito
        'Last Modification: 09/05/2020
        'Prepares the "CharacterAttackNpc" message and returns it.
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.CharacterAttackNpc)
102     Call Writer.WriteInt(charindex)

        '<EhFooter>
        Exit Function

PrepareMessageCharacterAttackNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCharacterAttackNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Writes the "StrDextRunningOut" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Seconds Seconds left.

Public Sub WriteStrDextRunningOut(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteStrDextRunningOut_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Dalmasso (CHOTS)
        'Last Modification: 08/06/2011
        '
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.StrDextRunningOut)
    
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteStrDextRunningOut_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStrDextRunningOut " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Handles the "PMSend" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePMSend(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePMSend_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Handles the "PMSend" message.
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName    As String

            Dim Message     As String

            Dim TargetIndex As Integer

102         UserName = Reader.ReadString8
104         Message = Reader.ReadString8
        
106         TargetIndex = NameIndex(UserName)
        
108         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
110             If TargetIndex = 0 Then 'Offline
112                 If FileExist(CharPath & UserName & ".chr", vbNormal) Then
114                     Call AgregarMensajeOFF(UserName, .Name, Message)
116                     Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)
                    Else
118                     Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else 'Online
120                 Call AgregarMensaje(TargetIndex, .Name, Message)
122                 Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandlePMSend_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePMSend " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleSearchObj(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSearchObj_Err
        '</EhHeader>

        '***************************************************
        'Author: WAICON
        'Last Modification: 06/05/2019
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Tag     As String

            Dim A       As Long

            Dim cant    As Long

            Dim strTemp As String
        
102         Tag = Reader.ReadString8
        
104         If .flags.Privilegios And (PlayerType.Admin) Then

106             For A = 1 To UBound(ObjData)

108                 If InStr(1, Tilde(ObjData(A).Name), Tilde(Tag)) Then
110                     strTemp = strTemp & A & " " & ObjData(A).Name & vbCrLf
                    
112                     cant = cant + 1
                    End If

                Next
            
114             If cant = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "No hubo resultados de: '" & Tag & "'", FontTypeNames.FONTTYPE_INFO)
                Else
118                 Call WriteConsoleMsg(UserIndex, "Hubo " & cant & " resultados de: " & Tag & strTemp, FontTypeNames.FONTTYPE_INFOBOLD)
                End If
        
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleSearchObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSearchObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleUserEditation(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleUserEditation_Err

        '</EhHeader>
    
100     With UserList(UserIndex)
        
            Dim Elv As Byte
            
            Select Case .Account.Premium
            
                Case 0
                    Exit Sub

                Case 1
                    Elv = 30

                Case 2
                    Elv = 35

                Case 3
                    Elv = 40

            End Select
            
104         If .Stats.Elv >= Elv Or .Stats.Elv < 3 Then
                Call WriteConsoleMsg(UserIndex, "Debes ser nivel inferior a " & Elv & " para poder reiniciar tu personaje de acuerdo al Tier elegido.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
        
106         If .GroupIndex > 0 Then
108             Call SaveExpAndGldMember(.GroupIndex, UserIndex)
110             Call WriteConsoleMsg(UserIndex, "Dirigete a una zona libre de party.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            If .flags.Navegando > 0 Then
                Call WriteConsoleMsg(UserIndex, "Deja de navegar y podrás reiniciar tu personaje.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            If MapInfo(.Pos.Map).Pk Then
                Call WriteConsoleMsg(UserIndex, "¡Vete a Zona Segura! Aquí corres peligro...", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            If .MascotaIndex > 0 Then
                Call QuitarPet(UserIndex, .MascotaIndex)

            End If
            
            ' Tiene objetos, los desequipamos
            Call Reset_DesquiparAll(UserIndex)

112         Call InitialUserStats(UserList(UserIndex))
            
            
            'Call QuitarNewbieObj(UserIndex)
            
            Call LimpiarInventario(UserIndex)
            Call ApplySetInitial_Newbie(UserIndex)
            Call UpdateUserInv(True, UserIndex, 0)
            
120         If MapInfo(.Pos.Map).LvlMin > .Stats.Elv Then
122             Call WriteConsoleMsg(UserIndex, "Hemos notado que no puedes sobrevivir a la peligrosidad de este mapa. ¡Serás llevado a Ullathorpe! ¡No nos lo agradezcas!", FontTypeNames.FONTTYPE_INFORED)
124             Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)

            End If
        
126         Call WriteLevelUp(UserIndex, .Stats.SkillPts)
128         Call WriteUpdateUserStats(UserIndex)
130         Call WriteUpdateStrenghtAndDexterity(UserIndex)
132         Call WriteConsoleMsg(UserIndex, "Has reiniciado tu personaje. ¡Que tengas un excelente re-comienzo!", FontTypeNames.FONTTYPE_GUILDMSG)

        End With
    
        '<EhFooter>
        Exit Sub

HandleUserEditation_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserEditation " & "at line " & Erl

        '</EhFooter>
End Sub

Private Sub HandlePartyClient(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePartyClient_Err
        '</EhHeader>

        ' 1) Requiere formulario 'principal'
        ' 4) Abandonar party
        ' 5) Requiere ingresar a party
    
        Dim Paso As Byte
    
100     With UserList(UserIndex)
        
102         Select Case Reader.ReadInt

                Case 1

104                 If .GroupIndex <= 0 Then
106                     mGroup.CreateGroup (UserIndex)
                    Else
108                     WriteGroupPrincipal (UserIndex)
                    End If
110             Case 2 ' Cambia la obtención de Experiencia, para ver si recibe por golpe o acumula...
112                 If .GroupIndex > 0 Then
114                     mGroup.ChangeObtainExp UserIndex
                
                    End If
                
116             Case 3
118                 mGroup.AcceptInvitationGroup UserIndex
120             Case 4

122                 If .GroupIndex > 0 Then
124                     mGroup.AbandonateGroup UserIndex
                    End If
            
126             Case 5
128                 mGroup.SendInvitationGroup UserIndex
            End Select

        End With

        '<EhFooter>
        Exit Sub

HandlePartyClient_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePartyClient " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteGroupUpdateExp(ByVal UserIndex As Integer, ByVal GroupIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteGroupUpdateExp_Err
        '</EhHeader>
100      Call Writer.WriteInt(ServerPacketID.GroupUpdateExp)
         Dim A As Long
     
102      With Groups(GroupIndex)
104         For A = 1 To MAX_MEMBERS_GROUP
106             Call Writer.WriteInt32(.User(A).Exp)
108         Next A
         End With
     
        '<EhFooter>
        Exit Sub

WriteGroupUpdateExp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteGroupUpdateExp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteGroupPrincipal(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteGroupPrincipal_Err
        '</EhHeader>

        Dim GroupIndex As Integer

        Dim A          As Long, B As Long
    
100     Call Writer.WriteInt(ServerPacketID.GroupPrincipal)
    
102     GroupIndex = UserList(UserIndex).GroupIndex
    
104     With Groups(GroupIndex)
106         Call Writer.WriteBool(.Acumular)
        
108         For A = 1 To MAX_MEMBERS_GROUP

110             If .User(A).Index > 0 Then
112                 Call Writer.WriteString8(UserList(.User(A).Index).Name)
                Else
114                 Call Writer.WriteString8("<Vacio>")
                End If
            
            
116             Call Writer.WriteInt8(.User(A).PorcExp)
118             Call Writer.WriteInt32(.User(A).Exp)
120         Next A

        End With
              
122     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteGroupPrincipal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteGroupPrincipal " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function CheckValidPorc(ByVal UserIndex As Integer, ByRef Exp() As Byte) As Boolean
        On Error GoTo CheckValidPorc_Err
    
        Dim A As Long
        Dim Porc As Long
    
        With UserList(UserIndex)
            If .Invent.PendientePartyObjIndex = 0 Then
                CheckValidPorc = False
                Exit Function
            End If
            
            Porc = ObjData(.Invent.PendientePartyObjIndex).Porc
            
             For A = LBound(Exp) To UBound(Exp)
                If Exp(A) > Porc Then
                    Exit Function
                End If
            Next A
        End With

        CheckValidPorc = True
        '<EhFooter>
        Exit Function

CheckValidPorc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.CheckValidPorc " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub HandleGroupChangePorc(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleGroupChangePorc_Err

        '</EhHeader>
    
        Dim A      As Byte

        Dim Exp(4) As Byte
    
100     With UserList(UserIndex)
        
102         For A = 0 To 4
104             Exp(A) = Reader.ReadInt
106         Next A
        
108         If .GroupIndex > 0 Then

                 If Not CheckValidPorc(UserIndex, Exp) Then
                    Call WriteConsoleMsg(UserIndex, "¡No tienes ningún Pendiente de Experiencia o bien no permite cambiar al porcentaje seleccionado!", FontTypeNames.FONTTYPE_ANGEL)
                    Exit Sub
                 End If

110             mGroup.GroupSetPorcentaje UserIndex, .GroupIndex, Exp

            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleGroupChangePorc_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGroupChangePorc " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub WriteUserInEvent(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserInEvent_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.UserInEvent)

102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUserInEvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserInEvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleEntrarDesafio(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEntrarDesafio_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
102         Select Case Reader.ReadInt
            
                Case 0
            
104                 Call mDesafios.Desafio_UserAdd(UserIndex)

106             Case 1
                    'If .flags.Desafiando Then
                    'Call Desafio_UserKill(UserIndex)
                    'End If
            End Select
    
        End With

        '<EhFooter>
        Exit Sub

HandleEntrarDesafio_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEntrarDesafio " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


''
' Writes the "MontateToggle" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMontateToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteMontateToggle_Err
        '</EhHeader>

        '***************************************************
        'Author: Dragons
        'Last Modification: 30/06/2019
        'Writes the "MontateToggle" message to the given user's outgoing data buffer
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.MontateToggle)

102     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteMontateToggle_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMontateToggle " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleSetPanelClient(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSetPanelClient_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
102         .flags.MenuCliente = Reader.ReadInt
104         .flags.LastSlotClient = Reader.ReadInt

              Dim X As Long, Y As Long
              
              X = Reader.ReadInt
              Y = Reader.ReadInt
              
              If Not (X = 0 And Y = 0) Then
106             UpdatePointer UserIndex, .flags.MenuCliente, X, Y, "Solapas Inv-Hec"
              End If
              
108         Reader.ReadInt16
        End With

        '<EhFooter>
        Exit Sub

HandleSetPanelClient_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSetPanelClient " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleSolicitaSeguridad(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSolicitaSeguridad_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim Tipo     As Byte

            Dim TempName As String, TempHD As String
        
102         UserName = Reader.ReadString8
104         Tipo = Reader.ReadInt
        
106         If CharIs_Admin(UCase$(.Name)) Then
            
108             tUser = NameIndex(UserName)
            
110             If tUser <= 0 Then
112                 WriteConsoleMsg UserIndex, "El usuario se ha desconectado.", FontTypeNames.FONTTYPE_INFO
                Else

114                 If EsGm(tUser) Then
116                     WriteConsoleMsg UserIndex, "No puedes ver la información de otros GameMaster", FontTypeNames.FONTTYPE_INFO
                    Else

118                     Select Case Tipo

                                ' Inicia el Seguimiento
                            Case 0

120                             If UserList(tUser).flags.GmSeguidor > 0 Then
122                                 WriteConsoleMsg UserList(tUser).flags.GmSeguidor, "El GM " & .Name & " ha comenzado a analizar al personaje " & UserList(tUser).Name, FontTypeNames.FONTTYPE_INFORED
124                                 UserList(tUser).flags.GmSeguidor = UserIndex
                                Else
126                                 UserList(tUser).flags.GmSeguidor = UserIndex
128                                 WriteSolicitaCapProc tUser, 0
                                End If
                            
130                             Call WriteUpdateListSecurity(UserList(tUser).flags.GmSeguidor, UserList(tUser).Name, vbNullString, 255)
                            
132                         Case 1 ' Actualiza la solapa de procesos
134                             WriteSolicitaCapProc tUser, 1
                            
136                         Case 2 ' Actualiza la solapa de captions
138                             WriteSolicitaCapProc tUser, 2
                        
140                         Case 3, 4, 5

142                             If .flags.Privilegios And (PlayerType.Admin) Then
144                                 WriteSolicitaCapProc tUser, Tipo
                                End If
                        
146                         Case Else

                        End Select
                    
                    End If
                End If
        
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleSolicitaSeguridad_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSolicitaSeguridad " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleSendListSecurity(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSendListSecurity_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim List As String

            Dim Tipo As Byte
        
102         List = Reader.ReadString8
104         Tipo = Reader.ReadInt
        
106         If .flags.GmSeguidor > 0 Then
108             Call WriteUpdateListSecurity(.flags.GmSeguidor, .Name, List, Tipo)
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleSendListSecurity_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSendListSecurity " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdateListSecurity(ByVal UserIndex As Integer, _
                                   ByVal CheaterName As String, _
                                   ByVal List As String, _
                                   ByVal Tipo As Byte)
        '<EhHeader>
        On Error GoTo WriteUpdateListSecurity_Err
        '</EhHeader>
    
100     Call Writer.WriteInt(ServerPacketID.UpdateListSecurity)
102     Call Writer.WriteString8(CheaterName)
104     Call Writer.WriteString8(List)
106     Call Writer.WriteInt(Tipo)
    
108     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateListSecurity_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateListSecurity " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteSolicitaCapProc(ByVal UserIndex As Integer, _
                                Optional ByVal Tipo As Byte = 0, _
                                Optional ByVal Process As String = vbNullString, _
                                Optional ByVal Captions As String = vbNullString)
        '<EhHeader>
        On Error GoTo WriteSolicitaCapProc_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.SolicitaCapProc)
102     Call Writer.WriteInt(Tipo)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteSolicitaCapProc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteSolicitaCapProc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function PrepareMessageCreateDamage(ByVal X As Byte, _
                                           ByVal Y As Byte, _
                                           ByVal DamageValue As Long, _
                                           ByVal DamageType As eDamageType, _
                                           Optional ByVal Text As String = vbNullString)
        '<EhHeader>
        On Error GoTo PrepareMessageCreateDamage_Err
        '</EhHeader>
 
100     Writer.WriteInt ServerPacketID.CreateDamage
102     Writer.WriteInt8 X
104     Writer.WriteInt8 Y
106     Writer.WriteInt32 DamageValue
108     Writer.WriteInt8 DamageType

110     If DamageType = d_AddMagicWord Then
112         Writer.WriteString8 Text
        End If
        '<EhFooter>
        Exit Function

PrepareMessageCreateDamage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCreateDamage " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function WriteVesA(ByVal UserIndex As Integer, _
                            ByVal Name As String, ByVal Desc As String, _
                            ByVal Class As eClass, ByVal Raza As eRaza, _
                            ByVal Faction As Byte, ByVal FactionRange As String, _
                            ByVal GuildName As String, ByVal GuildRange As Byte, _
                            ByVal RangeGm As String, ByVal sPlayerType As Byte, _
                            ByVal IsGold As Byte, ByVal IsBronce As Byte, _
                            ByVal IsPlata As Byte, ByVal IsPremium As Byte, _
                            ByVal IsStreamer As Byte, _
                            ByVal IsTransform As Byte, ByVal IsKilled As Byte, _
                            ByVal FtOptional As FontTypeNames, _
                            ByVal StreamerUrl As String, _
                            ByVal Rachas As Integer, _
                            ByVal RachasHist As Integer) As String
        '<EhHeader>
        On Error GoTo WriteVesA_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.ClickVesA)
        
102     Call Writer.WriteString8(Name)
104     Call Writer.WriteString8(Desc)
106     Call Writer.WriteInt(Class)
108     Call Writer.WriteInt(Raza)
110     Call Writer.WriteInt(Faction)
112     Call Writer.WriteString8(FactionRange)
114     Call Writer.WriteString8(GuildName)
116     Call Writer.WriteInt(GuildRange)
        
118     Call Writer.WriteString8(RangeGm)
120     Call Writer.WriteInt(sPlayerType)
        
122     Call Writer.WriteInt(IsGold)
124     Call Writer.WriteInt(IsBronce)
126     Call Writer.WriteInt(IsPlata)
128     Call Writer.WriteInt(IsPremium)
130     Call Writer.WriteInt(IsStreamer)
132     Call Writer.WriteInt(IsTransform)
134     Call Writer.WriteInt(IsKilled)
136     Call Writer.WriteInt(FtOptional)
138     Call Writer.WriteString8(StreamerUrl)
        Call Writer.WriteInt16(Rachas)
        Call Writer.WriteInt16(RachasHist)
140     Call SendData(ToOne, UserIndex, vbNullString)
   
        '<EhFooter>
        Exit Function

WriteVesA_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteVesA " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub HandleCheckingGlobal(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCheckingGlobal_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            
104             If GlobalActive Then
106                 GlobalActive = False
108                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat global ha sido desactivado.", FontTypeNames.FONTTYPE_INFO))
                Else
110                 GlobalActive = True
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat global ha sido activado. Utiliza el comando /GLOBAL para hablar con los demás usuarios del juego.", FontTypeNames.FONTTYPE_GUILD))
                End If

            End If

        End With

        '<EhFooter>
        Exit Sub

HandleCheckingGlobal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCheckingGlobal " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleChatGlobal(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChatGlobal_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim Message As String

102         Message = Reader.ReadString8
        
104         If .flags.Streamer = 1 Then
106             If GlobalActive = False Then
108                 Call WriteConsoleMsg(UserIndex, "El Chat Global se encuentra desactivado.", FontTypeNames.FONTTYPE_INFO)
110             ElseIf .Counters.TimeGlobal > 0 Then
112                 Call WriteConsoleMsg(UserIndex, "Debes esperar algunos segundos para volver a enviar un mensaje al global", FontTypeNames.FONTTYPE_INFO)
114             ElseIf .Counters.Pena > 0 Then
116                 Call WriteConsoleMsg(UserIndex, "No puedes enviar mensajes desde la cárcel", FontTypeNames.FONTTYPE_INFO)
118             ElseIf .flags.Silenciado > 0 Then
120                 Call WriteConsoleMsg(UserIndex, "Los administradores te han silenciado. No podrás enviar mensajes al Chat Global", FontTypeNames.FONTTYPE_INFO)
122             ElseIf Not AsciiValidos_Chat(Message) Then
124                 Call WriteConsoleMsg(UserIndex, "Mensaje inválido.", FontTypeNames.FONTTYPE_INFO)
                Else
126                 Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & "» " & Message, FontTypeNames.FONTTYPE_GLOBAL))
            
128                 .Counters.TimeGlobal = 3
            
                End If

            Else
130             Call WriteConsoleMsg(UserIndex, "Solo los personajes considerados STREAMERS OFICIALES pueden utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleChatGlobal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChatGlobal " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleCountDown(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleCountDown_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            Dim Count    As Byte

            Dim CountMap As Boolean
        
102         Count = Reader.ReadInt + 1
104         CountMap = Reader.ReadBool
        
106         If Count > 240 Then Exit Sub
        
108         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            
110             If Count = 0 Then
112                 CountDown_Map = 0
114                 CountDown_Time = 0

                    Exit Sub

                End If
            
116             If CountMap Then
118                 CountDown_Map = .Pos.Map
120                 CountDown_Time = Count
                Else
122                 CountDown_Time = Count
124                 CountDown_Map = 0
                End If
            
            End If
    
        End With

        '<EhFooter>
        Exit Sub

HandleCountDown_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleCountDown " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleGiveBackUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGiveBackUser_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8
        
104         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
106             tUser = NameIndex(UserName)
            
108             If tUser <= 0 Then
110                 WriteConsoleMsg UserIndex, "El usuario está offline.", FontTypeNames.FONTTYPE_INFO
                Else

112                 If UserList(tUser).PosAnt.Map <> 0 Then
114                     Call WarpPosAnt(tUser)
                    End If
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleGiveBackUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGiveBackUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleLearnMeditation(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLearnMeditation_Err
        '</EhHeader>
    
        Dim Tipo     As Byte

        Dim Selected As Byte
    
100     With UserList(UserIndex)
        
102         Tipo = Reader.ReadInt
104         Selected = Reader.ReadInt
        
106         If Selected < 0 Or Selected > MAX_MEDITATION Then Exit Sub
        
108         Select Case Tipo

                Case 0 ' Aprender nueva / Reclamar

110                 If Selected = 0 Then Exit Sub
                
112                 'Call mMeditations.Meditation_AddNew(UserIndex, Selected)
                
114             Case 1 ' Poner en uso
116                 Call mMeditations.Meditation_Select(UserIndex, Selected)
            End Select
    
        End With

        '<EhFooter>
        Exit Sub

HandleLearnMeditation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLearnMeditation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteCreateDamage(ByVal UserIndex As Integer, _
                             ByVal X As Byte, _
                             ByVal Y As Byte, _
                             ByVal Value As Long, _
                             ByVal DamageType As eDamageType)
        '<EhHeader>
        On Error GoTo WriteCreateDamage_Err
        '</EhHeader>

100     Call SendData(ToOne, UserIndex, PrepareMessageCreateDamage(X, Y, Value, DamageType))

        '<EhFooter>
        Exit Sub

WriteCreateDamage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteCreateDamage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleInfoEvento(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleInfoEvento_Err
        '</EhHeader>
    
        Dim A As Long
    
100     With UserList(UserIndex)
        
        If Not Interval_Packet250(UserIndex) Then Exit Sub
            
        Call WriteTournamentList(UserIndex)
        
        End With

        '<EhFooter>
        Exit Sub

HandleInfoEvento_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleInfoEvento " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub HandleDragToPos(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDragToPos_Err
        '</EhHeader>

        ' @ Author : maTih.-
        '            Drag&Drop de objetos en del inventario a una posición.
            
        Dim X      As Byte

        Dim Y      As Byte

        Dim Slot   As Byte

        Dim Amount As Integer

        Dim tUser  As Integer

        Dim tNpc   As Integer

100     X = Reader.ReadInt()
102     Y = Reader.ReadInt()
104     Slot = Reader.ReadInt()
106     Amount = Reader.ReadInt()

108     tUser = MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex

110     tNpc = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex

112     If Not Interval_Drop(UserIndex) Then Exit Sub
114     If UserList(UserIndex).flags.Comerciando Then Exit Sub
116     If UserList(UserIndex).flags.Montando Then Exit Sub
    
118     If Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Sub
120     If MapData(UserList(UserIndex).Pos.Map, X, Y).Blocked = 1 Then Exit Sub
122     If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
124     If Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount Then Exit Sub
126     If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then Exit Sub
128     If tUser = UserIndex Then Exit Sub
130     If EsGm(UserIndex) And Not EsGmPriv(UserIndex) Then Exit Sub
    
132     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).NoNada = 1 Then Exit Sub
    
        'If EsNewbie(UserIndex) Then
            'Call WriteConsoleMsg(UserIndex, "Los newbies no pueden dropear objetos.", FontTypeNames.FONTTYPE_INFO)
            'Exit Sub
        'End If
    
134     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otGemaTelep Then
136         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
138     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otMonturas Then
140         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
142     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otTransformVIP Then
144         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
146     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).NoDrop = 1 Then
148         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        
150     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Plata = 1 Then
152         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        
154     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Oro = 1 Then
156         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        
158     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Premium = 1 Then
160         Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
162     If UserList(UserIndex).flags.SlotEvent > 0 Or UserList(UserIndex).flags.SlotReto > 0 Then Exit Sub
    
164     If tUser > 0 Then
166         If tUser = UserIndex Then Exit Sub
168         If EsGm(tUser) Then Exit Sub
         
170         If UserList(tUser).flags.DragBlocked Then
172             Call WriteConsoleMsg(UserIndex, "La persona no quiere tus objetos.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
174         If UserList(tUser).flags.Comerciando Then
176             Call WriteConsoleMsg(UserIndex, "No puedes arrojar objetos si la persona está comerciando.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            
        
178         Call mDragAndDrop.DragToUser(UserIndex, tUser, Slot, Amount)

180     ElseIf tNpc > 0 Then
182         Call mDragAndDrop.DragToNPC(UserIndex, tNpc, Slot, Amount)

        Else
184         Call mDragAndDrop.DragToPos(UserIndex, X, Y, Slot, Amount)

        End If

        '<EhFooter>
        Exit Sub

HandleDragToPos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDragToPos " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleAbandonateFaction(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAbandonateFaction_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            'Validate target NPC
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

106         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

108         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
110             Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

112         If .Faction.Status = 0 Then
114             Call WriteConsoleMsg(UserIndex, "¡No perteneces a ninguna facción!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
116         Call mFacciones.Faction_RemoveUser(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

HandleAbandonateFaction_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAbandonateFaction " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleEnlist(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEnlist_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            'Validate target NPC
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

106         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

108         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
110             Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

112         If .Faction.Status > 0 Then
114             Call WriteConsoleMsg(UserIndex, "Ya eres miembro de una facción y espero que sea la nuestra, sino mis guardias te atacaran!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
116         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
118             If Escriminal(UserIndex) Then
120                 Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                    Exit Sub

                End If
            
122             Call mFacciones.Faction_AddUser(UserIndex, r_Armada)
124             Call Guilds_CheckAlineation(UserIndex, a_Armada)
            Else

126             If Not Escriminal(UserIndex) Then
128                 Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                    Exit Sub

                End If
            
130             Call mFacciones.Faction_AddUser(UserIndex, r_Caos)
132             Call Guilds_CheckAlineation(UserIndex, a_Legion)
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleEnlist_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEnlist " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleReward(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleReward_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            'Validate target NPC
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

106         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

108         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
110             Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
    
112         Call mFacciones.Faction_CheckRangeUser(UserIndex)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleReward_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleReward " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleFianza(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleFianza_Err
        '</EhHeader>
    
        '***************************************************
        'Author: Matías Ezequiel
        'Last Modification: 16/03/2016 by DS
        'Sistema de fianzas TDS.
        '***************************************************
        Dim Fianza As Long
        Dim ValueFianza As Long
        
100     With UserList(UserIndex)
        
102         Fianza = Reader.ReadInt
            ValueFianza = Fianza * 5
            
104         If Fianza <= 0 Or Fianza > MAXORO Then Exit Sub
        
106         If MapInfo(.Pos.Map).Pk Then
108             Call WriteConsoleMsg(UserIndex, "Debes estar en zona segura para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
110         If .flags.Muerto Then
112             Call WriteConsoleMsg(UserIndex, "Estás muerto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

114         If (ValueFianza) > .Stats.Gld Then
116             Call WriteConsoleMsg(UserIndex, "Para pagar esa fianza necesitas pagar impuestos. El precio total es: " & (ValueFianza) & " Monedas de Oro. ¡Agradece que no son Eldhires gusano!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            Dim EraCriminal As Boolean
        
118         EraCriminal = Escriminal(UserIndex)
120         .Reputacion.NobleRep = .Reputacion.NobleRep + Fianza
122         .Stats.Gld = .Stats.Gld - (ValueFianza)

124         Call WriteConsoleMsg(UserIndex, "Has ganado " & Fianza & " puntos de noble.", FontTypeNames.FONTTYPE_INFO)
126         Call WriteConsoleMsg(UserIndex, "Se te han descontado " & ValueFianza & " Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)
128         Call WriteUpdateGold(UserIndex)
        
130         If EraCriminal And Not Escriminal(UserIndex) Then
132             Call RefreshCharStatus(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleFianza_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleFianza " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleHome(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleHome_Err

        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Creation Date: 06/01/2010
        'Last Modification: 05/06/10
        'Pato - 05/06/10: Add the Ucase$ to prevent problems.
        '***************************************************
100     With UserList(UserIndex)
            
            
            ' @ El personaje se asocia a una nueva CIUDAD.
            If .flags.TargetNPC > 0 Then
                If Npclist(.flags.TargetNPC).Ciudad > 0 Then
                    Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
                    Exit Sub
                End If
            End If
            
102         If .flags.Muerto = 0 Then
104             Call WriteConsoleMsg(UserIndex, "No puedes usar el comando si estás vivo.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
106         If Not MapInfo(.Pos.Map).Pk Then
108             Call WriteConsoleMsg(UserIndex, "Ya te encuentras en Zona Segura", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If .flags.Traveling = 1 Then
                Call EndTravel(UserIndex, True)
                Exit Sub

            End If
                
            Dim RequiredGld As Long
            
            If Not EsNewbie(UserIndex) Then
110             If .Stats.Elv < 20 Then
112                 RequiredGld = 10 * .Stats.Elv
114             ElseIf .Stats.Elv < 35 Then
116                 RequiredGld = 50 * .Stats.Elv
                Else
118                 RequiredGld = 150 * .Stats.Elv

                End If

            End If

            
            Select Case .Account.Premium
                Case 0
                
                Case 1, 2
                    RequiredGld = RequiredGld / 2
                Case 3
                    RequiredGld = 0
            End Select
            
120
            
122         If .flags.SlotEvent > 0 Then
124             WriteConsoleMsg UserIndex, "No puedes usar la restauración si estás en un evento.", FontTypeNames.FONTTYPE_INFO

                Exit Sub

            End If
              
126         If .flags.SlotReto > 0 Or .flags.SlotFast > 0 Then
128             WriteConsoleMsg UserIndex, "No puede susar este comando si estás en reto.", FontTypeNames.FONTTYPE_INFO

                Exit Sub

            End If
        
130         If .Counters.Pena > 0 Then
132             Call WriteConsoleMsg(UserIndex, "No puedes usar la restauración si estás en la carcel.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
134         If .Stats.Gld < RequiredGld Then
136             Call WriteConsoleMsg(UserIndex, "El viaje requiere que dispongas de " & RequiredGld & " Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            .Stats.Gld = .Stats.Gld - RequiredGld
            Call WriteUpdateGold(UserIndex)
            
            Call goHome(UserIndex)

146

        End With

        '<EhFooter>
        Exit Sub

HandleHome_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleHome " & "at line " & Erl

        

        '</EhFooter>
End Sub

''
' Prepares the "UpdateControlPotas" message and returns it.
'

Public Function PrepareMessageUpdateControlPotas(ByVal charindex As Integer, _
                                                 ByVal MinHp As Integer, _
                                                 ByVal MaxHp As Integer, _
                                                 ByVal MinMan As Integer, _
                                                 ByVal MaxMan As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateControlPotas_Err
        '</EhHeader>

        '***************************************************
        'Author
        'Last Modification:
        '
        '***************************************************
100     Call Writer.WriteInt(ServerPacketID.UpdateControlPotas)
        
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(MinHp)
106     Call Writer.WriteInt(MaxHp)
108     Call Writer.WriteInt(MinMan)
110     Call Writer.WriteInt(MaxMan)

        '<EhFooter>
        Exit Function

PrepareMessageUpdateControlPotas_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageUpdateControlPotas " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

'
' Handles the "SendReply" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSendReply(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSendReply_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim Fight As tFight
            Dim Users() As String
            Dim Temp As String
            Dim A As Long
            

102         Temp = Reader.ReadString8
            Fight.Tipo = Reader.ReadInt8
104         Fight.Gld = Reader.ReadInt32
108         Fight.Time = (Reader.ReadInt8 * 60)
110         Fight.RoundsLimit = Reader.ReadInt8
112         Fight.Terreno = Reader.ReadInt8

114         For A = LBound(Fight.config) To UBound(Fight.config)
116             Fight.config(A) = Reader.ReadInt8
118         Next A
    
120         Users = Split(Temp, "-")
        
        
122         ReDim Fight.User(LBound(Users) To UBound(Users)) As tFightUser
        
124         For A = LBound(Users) To UBound(Users)
126             Fight.User(A).Name = Users(A)
128         Next A
        
130         Call mRetos.SendFight(UserIndex, Fight)
              
        End With
    
        '<EhFooter>
        Exit Sub

HandleSendReply_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSendReply " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

'
' Handles the "AcceptReply" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptReply(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAcceptReply_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim UserName As String
        
102         UserName = Reader.ReadString8
                      
104         Call mRetos.AcceptFight(UserIndex, UserName)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleAcceptReply_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAcceptReply " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

'
' Handles the "AbandonateReply" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAbandonateReply(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAbandonateReply_Err
        '</EhHeader>

        '***************************************************
        'Author:
        'Last Modification:
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If .flags.SlotReto > 0 Then
104             Call mRetos.UserdieFight(UserIndex, 0, True)
            End If
    
        End With

        '<EhFooter>
        Exit Sub

HandleAbandonateReply_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAbandonateReply " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleEvents_CreateNew(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEvents_CreateNew_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim Temp As tEvents
        
            Dim A                  As Integer
        
102         Temp.Modality = Reader.ReadInt
104         Temp.Name = Reader.ReadString8
              Temp.QuotasMin = Reader.ReadInt
106         Temp.QuotasMax = Reader.ReadInt
108         Temp.LvlMin = Reader.ReadInt
110         Temp.LvlMax = Reader.ReadInt
112         Temp.InscriptionGld = Reader.ReadInt
114         Temp.InscriptionEldhir = Reader.ReadInt
116         Temp.TimeInscription = Reader.ReadInt
118         Temp.TimeCancel = Reader.ReadInt
120         Temp.TeamCant = Reader.ReadInt
                      
122         Temp.LimitRed = Reader.ReadInt
124         Temp.PrizeGld = Reader.ReadInt
126         Temp.PrizeEldhir = Reader.ReadInt
128         Temp.PrizeObj.ObjIndex = Reader.ReadInt
130         Temp.PrizeObj.Amount = Reader.ReadInt
        
132         ReDim Temp.AllowedClasses(1 To NUMCLASES) As Byte
134         ReDim Temp.AllowedFaction(1 To 4) As Byte
        
136         For A = 1 To 4
138             Temp.AllowedFaction(A) = Reader.ReadInt()
140         Next A

142         For A = 1 To NUMCLASES
144             Temp.AllowedClasses(A) = Reader.ReadInt()
146         Next A
                      
148         Temp.ChangeClass = Reader.ReadInt()
150         Temp.ChangeRaze = Reader.ReadInt()
        
152         For A = 1 To MAX_EVENTS_CONFIG
154             Temp.config(A) = Reader.ReadInt8()
156         Next A
        
158         Temp.LimitRound = Reader.ReadInt8()
              Temp.LimitRoundFinal = Reader.ReadInt8()
160         Temp.GanaSigue = Reader.ReadInt8()
              Temp.ArenasLimit = Reader.ReadInt8()
              Temp.ArenasMin = Reader.ReadInt8()
              Temp.ArenasMax = Reader.ReadInt()
              Temp.ChangeLevel = Reader.ReadInt8()
              Temp.Prob = Reader.ReadInt8()
              
        
164         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
166             Dim CanEvent As Byte: CanEvent = NewEvent(Temp)
            
168             If CanEvent <> 0 Then
170                 Events(CanEvent).Enabled = True
172
                Else
174                 Call WriteConsoleMsg(UserIndex, "No hay más cupos para eventos o bien ya existe un evento con esa modalidad", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleEvents_CreateNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEvents_CreateNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleEvents_Close(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEvents_Close_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            Dim Slot As Byte
        
102         Slot = Reader.ReadInt
        
104         If Slot <= 0 Or Slot > MAX_EVENT_SIMULTANEO Then Exit Sub
        
106         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            
112             Call EventosDS.CloseEvent(Slot, , True)
            End If
        
        End With

        '<EhFooter>
        Exit Sub

HandleEvents_Close_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEvents_Close " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Private Sub HandlePro_Seguimiento(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandlePro_Seguimiento_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer

            Dim Seguir   As Boolean
        
102         UserName = Reader.ReadString8
104         Seguir = Reader.ReadBool
        
106         tUser = NameIndex(UserName)
        
108         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
110             If tUser > 0 Then
112                 If Seguir Then
114                     UserList(tUser).flags.GmSeguidor = UserIndex
116                     Call WriteConsoleMsg(UserIndex, "Has comenzado el seguimiento al usuario " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_INFOGREEN)
                    Else
118                     UserList(tUser).flags.GmSeguidor = 0
120                     Call WriteConsoleMsg(UserIndex, "Has reiniciado el seguimiento al usuario " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_INFOGREEN)
                    End If

                Else
122                 Call WriteConsoleMsg(UserIndex, "El personaje está offline", FontTypeNames.FONTTYPE_INFORED)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandlePro_Seguimiento_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandlePro_Seguimiento " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Function PrepareMessageUpdateGroupIndex(ByVal charindex As Integer, _
                                               ByVal GroupIndex As Byte) As String
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateGroupIndex_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.UpdateGroupIndex)
102     Call Writer.WriteInt(charindex)
104     Call Writer.WriteInt(GroupIndex)

        '<EhFooter>
        Exit Function

PrepareMessageUpdateGroupIndex_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageUpdateGroupIndex " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub WriteUpdateInfoIntervals(ByVal UserIndex As Integer, _
                                    ByVal Tipo As Byte, _
                                    ByVal Value As Long, _
                                    ByVal MenuCliente As Byte)
        '<EhHeader>
        On Error GoTo WriteUpdateInfoIntervals_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.UpdateInfoIntervals)
        
102     Call Writer.WriteInt(Tipo)
104     Call Writer.WriteInt(Value)
106     Call Writer.WriteInt(MenuCliente)
        
108     Call SendData(ToOne, UserIndex, vbNullString)
    
        '<EhFooter>
        Exit Sub

WriteUpdateInfoIntervals_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateInfoIntervals " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Private Sub HandleEvent_Participe(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEvent_Participe_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim Modality As String
            Dim Slot As Byte
            Dim ErrorMsg As String
        
102         Modality = Reader.ReadString8
104         Slot = Events_SearchSlotEvent(Modality): If Slot = 0 Then Exit Sub
    
            If Not Event_CheckInscriptions_User(UserIndex, Slot, ErrorMsg) Then
                Call WriteConsoleMsg(UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
            
            If Event_CheckExistUser(UserIndex, Slot) > 0 Then
                Call WriteConsoleMsg(UserIndex, "Ya estás inscripto en esta partida. Espera a que se complete y participarás automáticamente.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
        
            If .GroupIndex > 0 Then
                Call Events_Group_Set(.GroupIndex, Slot)
            Else
                Call Event_SetNewUser(UserIndex, Slot)
            End If
            
            
        End With
    
        '<EhFooter>
        Exit Sub

HandleEvent_Participe_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEvent_Participe " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Private Sub HandleUpdateInactive(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleUpdateInactive_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
102         .Counters.TimeInactive = 0
    
        End With

        '<EhFooter>
        Exit Sub

HandleUpdateInactive_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleUpdateInactive " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleRetos_RewardObj(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRetos_RewardObj_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
         
102         If .flags.ClainObject = 0 Then Exit Sub
104         If MapInfo(.Pos.Map).Pk Then Exit Sub
        
106         Call Retos_ReclameObj(UserIndex)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleRetos_RewardObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRetos_RewardObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleEvents_KickUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEvents_KickUser_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim tUser    As Integer
        
102         UserName = Reader.ReadString8
        
104         tUser = NameIndex(UserName)
        
106         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
108             If tUser > 0 Then
110                 Call EventosDS.AbandonateEvent(tUser)
112                 Call WriteConsoleMsg(UserIndex, "Has kickeado del evento al personaje " & UserList(tUser).Name & ". ¡No abuses de tu poder!", FontTypeNames.FONTTYPE_INFOGREEN)
114                 Call Logs_User(.Name, eLog.eGm, eNone, "El GM " & .Name & " ha kickeado del evento al personaje " & UserList(tUser).Name & ".")
                Else
116                 Call WriteConsoleMsg(UserIndex, "El personaje está offline. Actualiza la lista de eventos.", FontTypeNames.FONTTYPE_INFORED)
                End If
            End If
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleEvents_KickUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEvents_KickUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Required(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGuilds_Required_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            Dim Value As Integer

102         Value = Reader.ReadInt
        
104         If Value = 0 Then
106             Call WriteGuild_List(UserIndex, Guilds_PrepareList)
108         ElseIf Value > 0 And Value < MAX_GUILDS Then
110             Call WriteGuild_Info(UserIndex, Value, GuildsInfo(Value), GuildsInfo(Value).Members)
            Else
            
112             Select Case Value
            
                    Case 1000
114                     Call Guilds_PrepareInfoUsers(UserIndex)
                End Select
            
            End If
    
        End With

        '<EhFooter>
        Exit Sub

HandleGuilds_Required_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGuilds_Required " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteGuild_List(ByVal UserIndex As Integer, ByRef GuildList() As String)
        '<EhHeader>
        On Error GoTo WriteGuild_List_Err
        '</EhHeader>

        Dim Tmp As String

        Dim A   As Long
    
100     Call Writer.WriteInt(ServerPacketID.Guild_List)
102     Call Writer.WriteBool(UserList(UserIndex).GuildRange = rLeader Or UserList(UserIndex).GuildRange = rFound)
        
104     For A = LBound(GuildList()) To UBound(GuildList())
106         Tmp = Tmp & GuildList(A) & SEPARATOR
108     Next A
        
110     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
112     Call Writer.WriteString8(Tmp)
    
114     For A = 1 To MAX_GUILDS
116         Call Writer.WriteInt8(GuildsInfo(A).Alineation)
118         Call Writer.WriteInt8(GuildsInfo(A).NumMembers)
120         Call Writer.WriteInt8(GuildsInfo(A).MaxMembers)
122         Call Writer.WriteInt8(GuildsInfo(A).Lvl)
124         Call Writer.WriteInt32(GuildsInfo(A).Exp)
126         Call Writer.WriteInt32(GuildsInfo(A).Elu)
128     Next A
    
130     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteGuild_List_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteGuild_List " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Found(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGuilds_Found_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim Name                        As String, Temp As String

            Dim Alineation                  As eGuildAlineation

            Dim Codex(1 To MAX_GUILD_CODEX) As String

            Dim A                           As Long
        
102         Name = Reader.ReadString8
104         Alineation = Reader.ReadInt8
        
106         Call mGuilds.Guilds_New(UserIndex, Name, Alineation, Codex)
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleGuilds_Found_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGuilds_Found " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Invitation(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGuilds_Invitation_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            Dim UserName As String

            Dim Tipo     As Byte
        
102         UserName = Reader.ReadString8
104         Tipo = Reader.ReadInt
        
106         Select Case Tipo
        
                Case 0  ' El lider enviá solicitud a un miembro.
108                 Call Guilds_SendInvitation(UserIndex, UserName)

110             Case 1 ' El personaje acepta la solicitud del Lider
112                 Call Guilds_AcceptInvitation(UserIndex, UserName)
            End Select
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleGuilds_Invitation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGuilds_Invitation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteGuild_Info(ByVal UserIndex As Integer, _
                           ByVal GuildIndex As Integer, _
                           ByRef GuildInfo As tGuild, _
                           ByRef MemberInfo() As tGuildMember)
        '<EhHeader>
        On Error GoTo WriteGuild_Info_Err
        '</EhHeader>

        Dim A As Long
    
100     Call Writer.WriteInt(ServerPacketID.Guild_Info)
102     Call Writer.WriteInt16(GuildIndex)
    
104     Call Writer.WriteString8(GuildInfo.Name)
106     Call Writer.WriteInt8(GuildInfo.Alineation)
    
108     For A = 1 To MAX_GUILD_MEMBER
110         Call Writer.WriteString8(MemberInfo(A).Name)
112         Call Writer.WriteInt8(MemberInfo(A).Range)
            
114         Call Writer.WriteInt16(MemberInfo(A).Char.Body)
116         Call Writer.WriteInt16(MemberInfo(A).Char.Head)
118         Call Writer.WriteInt16(MemberInfo(A).Char.Helm)
120         Call Writer.WriteInt16(MemberInfo(A).Char.Shield)
122         Call Writer.WriteInt16(MemberInfo(A).Char.Weapon)
            
124     Next A
        
126     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteGuild_Info_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteGuild_Info " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Online(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGuilds_Online_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
102         If .GuildIndex = 0 Then
104             Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFORED)
            Else
106             Call Guilds_PrepareOnline(UserIndex, .GuildIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

HandleGuilds_Online_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGuilds_Online " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteGuild_InfoUsers(ByVal UserIndex As Integer, _
                                ByVal GuildIndex As Integer, _
                                ByRef MemberInfo() As tGuildMember)
        '<EhHeader>
        On Error GoTo WriteGuild_InfoUsers_Err
        '</EhHeader>

        Dim A As Long
    
100     Call Writer.WriteInt(ServerPacketID.Guild_InfoUsers)

102     Call Writer.WriteInt16(GuildIndex)
    
104     For A = 1 To MAX_GUILD_MEMBER
106         Call Writer.WriteString8(MemberInfo(A).Name)
108         Call Writer.WriteInt(MemberInfo(A).Range)
            
110         Call Writer.WriteInt(MemberInfo(A).Char.Elv)
112         Call Writer.WriteInt(MemberInfo(A).Char.Class)
114         Call Writer.WriteInt(MemberInfo(A).Char.Raze)
            
116         Call Writer.WriteInt(MemberInfo(A).Char.Body)
118         Call Writer.WriteInt(MemberInfo(A).Char.Head)
120         Call Writer.WriteInt(MemberInfo(A).Char.Helm)
122         Call Writer.WriteInt(MemberInfo(A).Char.Shield)
124         Call Writer.WriteInt(MemberInfo(A).Char.Weapon)
126         Call Writer.WriteInt(MemberInfo(A).Char.Points)
            
128     Next A
        
130     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteGuild_InfoUsers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteGuild_InfoUsers " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Kick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGuilds_Kick_Err
        '</EhHeader>
    
        Dim UserName As String

100     With UserList(UserIndex)
        
102         UserName = Reader.ReadString8

104         Call Guilds_KickUser(UserIndex, UCase$(UserName))
        End With
    
        '<EhFooter>
        Exit Sub

HandleGuilds_Kick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGuilds_Kick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Abandonate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGuilds_Abandonate_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
102         If .GuildIndex > 0 Then
104             Call Guilds_KickMe(UserIndex)
            Else
106             Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFORED)
            End If

        End With
    
        '<EhFooter>
        Exit Sub

HandleGuilds_Abandonate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGuilds_Abandonate " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Public Sub WriteFight_PanelAccept(ByVal UserIndex As Integer, _
                                  ByVal UserName As String, _
                                  ByVal TextUsers As String, _
                                  ByRef RetoTemp As tFight)
        '<EhHeader>
        On Error GoTo WriteFight_PanelAccept_Err
        '</EhHeader>

        Dim A   As Long

        Dim Str As String
        Dim Temp As Byte
    
100     Call Writer.WriteInt(ServerPacketID.Fight_PanelAccept)
        
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(TextUsers)
106     Call Writer.WriteInt32(RetoTemp.Gld)
110     Call Writer.WriteInt8(RetoTemp.RoundsLimit)
112     Call Writer.WriteInt8(RetoTemp.Terreno)
    
114     Temp = Int(RetoTemp.Time / 60)
116     Call Writer.WriteInt8(Temp)
    
118     For A = 1 To MAX_RETOS_CONFIG
120         Call Writer.WriteInt8(RetoTemp.config(A))
122     Next A
        
124     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteFight_PanelAccept_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteFight_PanelAccept " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleFight_CancelInvitation(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleFight_CancelInvitation_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         .Counters.FightInvitation = 0

        End With

        '<EhFooter>
        Exit Sub

HandleFight_CancelInvitation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleFight_CancelInvitation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleGuilds_Talk(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleGuilds_Talk_Err

        '</EhHeader>

100     With UserList(UserIndex)

            Dim chat      As String

            Dim IsSupport As Boolean
            
            Dim CanTalk   As Boolean
              
102         chat = Reader.ReadString8()
            IsSupport = Reader.ReadBool()
              
104         If LenB(chat) <> 0 Then
                  
106             CanTalk = True

108             If .flags.SlotEvent > 0 Then
110                 If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
112                     CanTalk = False

                    End If

                End If
                  
114             If CanTalk Then
116                 If .GuildIndex > 0 Then
                        If IsSupport Then
                            If GuildsInfo(.GuildIndex).Lvl >= 3 Then
                                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("[AYUDA] " & .Name & "> " & chat, FontTypeNames.FONTTYPE_INFORED))
                            End If
                        Else
                             Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("[CLANES]" & .Name & "> " & chat, FontTypeNames.FONTTYPE_GUILDMSG))
                        End If

118
                    
                    End If

                End If

            End If

        End With
    
        '<EhFooter>
        Exit Sub

HandleGuilds_Talk_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Talk " & "at line " & Erl

        

        '</EhFooter>
End Sub


Sub LoadPictureConMatrizDeBytes(ByRef Arrai() As Byte)
        '<EhHeader>
        On Error GoTo LoadPictureConMatrizDeBytes_Err
        '</EhHeader>

100     Dim FilePath  As String: FilePath = App.Path & "PRUEBA.BMP"
  
        Dim FileIndex As Integer

        Dim A         As Long
    
102     FileIndex = FreeFile
  
104     Open FilePath For Output As FileIndex

106     For A = LBound(Arrai) To UBound(Arrai)
108         Print #FileIndex, Arrai(A)
110     Next A

112     Close FileIndex
  
        '<EhFooter>
        Exit Sub

LoadPictureConMatrizDeBytes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.LoadPictureConMatrizDeBytes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
 
Private Sub HandleSendPic(ByVal UserIndex As Integer)
   
End Sub

Public Sub HandleLoginAccount(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLoginAccount_Err
        '</EhHeader>


        
        Dim Version As String, Email As String, Passwd As String
        Dim Time As Long
        Dim SERIAL(7) As String
        Dim Temp As tAccountSecurity
        'Dim Key_Encrypt As String: Key_Encrypt = mEncrypt_B.XOREncryption("ILMWNlOOvtUkOjo6bu")
        'Dim Key_Decrypt As String: Key_Decrypt = "ILMWNlOOvtUkOjo6bu"
    
100     Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    
    
        'Email = mEncrypt_A.AesDecryptString(Reader.ReadString8, mEncrypt_B.XOR_CHARACTER)
        'Passwd = mEncrypt_A.AesDecryptString(Reader.ReadString8, mEncrypt_B.XOR_CHARACTER)
102     Email = Reader.ReadString8
104     Passwd = Reader.ReadString8
106     Temp.SERIAL_BIOS = Reader.ReadString8 ' Serial_Bios
108     Temp.SERIAL_DISK = Reader.ReadString8 ' Serial_DISK
110     Temp.SERIAL_MAC = Reader.ReadString8 ' Serial_MAC
112     Temp.SERIAL_MOTHERBOARD = Reader.ReadString8 ' Serial_MOTHERBOARD
114     Temp.SERIAL_PROCESSOR = Reader.ReadString8 ' Serial_PROCESSOR
116     Temp.SYSTEM_DATA = Reader.ReadString8 ' System Data
118     Temp.IP_Local = Reader.ReadString8 ' IP LOCAL
120     Temp.IP_Public = Reader.ReadString8 ' IP PUBLICA

    
        Time = GetTime
        
        'If SERIAL(0) <> vbNullString Then Temp.SERIAL_BIOS = mEncrypt_A.AesDecryptString(SERIAL(0), Key_Decrypt)
        'If SERIAL(1) <> vbNullString Then Temp.SERIAL_DISK = mEncrypt_A.AesDecryptString(SERIAL(1), Key_Encrypt)
        'If SERIAL(2) <> vbNullString Then Temp.SERIAL_MAC = mEncrypt_A.AesDecryptString(SERIAL(2), Key_Decrypt)
        'If SERIAL(3) <> vbNullString Then Temp.SERIAL_MOTHERBOARD = mEncrypt_A.AesDecryptString(SERIAL(3), Key_Encrypt)
        'If SERIAL(4) <> vbNullString Then Temp.SERIAL_PROCESSOR = mEncrypt_A.AesDecryptString(SERIAL(4), Key_Decrypt)
        'If SERIAL(5) <> vbNullString Then Temp.SYSTEM_DATA = mEncrypt_A.AesDecryptString(SERIAL(5), Key_Encrypt)
        'If SERIAL(6) <> vbNullString Then Temp.IpAddress_Local = mEncrypt_A.AesDecryptString(SERIAL(6), Key_Encrypt)
        'If SERIAL(7) <> vbNullString Then Temp.IpAddress_Public = mEncrypt_A.AesDecryptString(SERIAL(7), Key_Decrypt)
    
        Email = LCase$(Email)
        
        Dim Testing As Boolean
        Const TIMER_MS As Byte = 250

        If (Time - TIMER_MS) <= TIMER_MS Then Exit Sub
        
        #If Testeo = 1 Then
122         Testing = True
        #End If


        
124     If SLOT_TERMINAL_ARCHIVE = 0 And Not Testing Then
126         Call Protocol.Kick(UserIndex, "Servidor en mantenimiento. Consulta otros servidores para disfrutar y pasar el rato.")
        
        Else
128         If ServerSoloGMs > 0 Then
130               If Not Email_Is_Testing_Pro(Email) Then
132                 Call Protocol.Kick(UserIndex, "Servidor en mantenimiento. Consulta otros servidores para disfrutar y pasar el rato.")
                    Exit Sub
        
                End If
            End If
    
134         If Not VersionOK(Version) Then
136             Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
            Else
138             If mAccount.LoginAccount(UserIndex, LCase$(Email), Passwd) Then
140                 UserList(UserIndex).Account.Sec = Temp
142                 UserList(UserIndex).Account.Sec.IP_Address = UserList(UserIndex).IpAddress
                    'UserList(UserIndex).IpAddress = UserList(UserIndex).Account.Sec.IP_Public
144                 Call Logs_Account_SettingData(UserIndex, "LOGIN", LCase$(Email))
                    Call WriteRequestID(LCase$(Email))
                End If
            End If
        End If
        
        
        UserList(UserIndex).LastRequestLogin = Time

        '<EhFooter>
        Exit Sub

HandleLoginAccount_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLoginAccount " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' Actualiza Slot de Mercado y/o Monedas de Oro de la Cuenta, como así también PREMIUM.
Public Sub WriteAccountInfo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteAccountInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.AccountInfo)
    
102     Call Writer.WriteInt32(UserList(UserIndex).Account.Gld)
104     Call Writer.WriteInt32(UserList(UserIndex).Account.Eldhir)
106     Call Writer.WriteInt8(UserList(UserIndex).Account.Premium)
108     Call Writer.WriteInt16(UserList(UserIndex).Account.MercaderSlot)
110     Call Writer.WriteInt32(UserList(UserIndex).Stats.Points)
112     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteAccountInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteAccountInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteLoggedAccount(ByVal UserIndex As Integer, ByRef Temp() As tAccountChar)
        '<EhHeader>
        On Error GoTo WriteLoggedAccount_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.loggedaccount)
    
        Dim A As Long
    
102     Call Writer.WriteInt32(UserList(UserIndex).Account.Gld)
104     Call Writer.WriteInt32(UserList(UserIndex).Account.Eldhir)
106     Call Writer.WriteInt32(UserList(UserIndex).Stats.Points)
    
108     Call Writer.WriteInt8(UserList(UserIndex).Account.Premium)
110     Call Writer.WriteInt16(UserList(UserIndex).Account.MercaderSlot)
    
112     Call Writer.WriteInt8(UserList(UserIndex).Account.CharsAmount)
    
114     For A = 1 To ACCOUNT_MAX_CHARS
118             Call Writer.WriteInt8(A)
120             Call Writer.WriteString8(Temp(A).Name)
122             Call Writer.WriteInt8(Temp(A).Blocked)
124             Call Writer.WriteString8(Temp(A).Guild)
            
126             Call Writer.WriteInt16(Temp(A).Body)
128             Call Writer.WriteInt16(Temp(A).Head)
130             Call Writer.WriteInt16(Temp(A).Weapon)
132             Call Writer.WriteInt16(Temp(A).Shield)
134             Call Writer.WriteInt16(Temp(A).Helm)
            
136             Call Writer.WriteInt8(Temp(A).Ban)
            
138             Call Writer.WriteInt8(Temp(A).Class)
140             Call Writer.WriteInt8(Temp(A).Raze)
142             Call Writer.WriteInt8(Temp(A).Elv)
            
144             Call Writer.WriteInt16(Temp(A).Map)
146             Call Writer.WriteInt8(Temp(A).posX)
148             Call Writer.WriteInt8(Temp(A).posY)
            
150             Call Writer.WriteInt8(Temp(A).Faction)
152             Call Writer.WriteInt8(Temp(A).FactionRange)
154     Next A
    
    
    
156     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteLoggedAccount_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteLoggedAccount " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteLoggedAccount_DataChar(ByVal UserIndex As Integer, _
                                       ByVal Slot As Byte, _
                                       ByRef DataChar As tAccountChar)
        '<EhHeader>
        On Error GoTo WriteLoggedAccount_DataChar_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.LoggedAccount_DataChar)


102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteString8(DataChar.Name)
106     Call Writer.WriteString8(DataChar.Guild)
        
108     Call Writer.WriteInt16(DataChar.Body)
110     Call Writer.WriteInt16(DataChar.Head)
112     Call Writer.WriteInt16(DataChar.Weapon)
114     Call Writer.WriteInt16(DataChar.Shield)
116     Call Writer.WriteInt16(DataChar.Helm)
        
118     Call Writer.WriteInt8(DataChar.Ban)
        
120     Call Writer.WriteInt8(DataChar.Class)
122     Call Writer.WriteInt8(DataChar.Raze)
124     Call Writer.WriteInt8(DataChar.Elv)
        
126     Call Writer.WriteInt16(DataChar.Map)
128     Call Writer.WriteInt8(DataChar.posX)
130     Call Writer.WriteInt8(DataChar.posY)
        
132     Call Writer.WriteInt8(DataChar.Faction)
134     Call Writer.WriteInt8(DataChar.FactionRange)
    
136     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteLoggedAccount_DataChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteLoggedAccount_DataChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteConnectedMessage(ByVal UserIndex As Integer, ByVal ServerSelected As Byte)
        '<EhHeader>
        On Error GoTo WriteConnectedMessage_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.Connected)
102     Call Writer.WriteInt8(ServerSelected)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteConnectedMessage_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteConnectedMessage " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteLoggedRemoveChar(ByVal UserIndex As Integer, ByVal SlotUserName As Byte)
        '<EhHeader>
        On Error GoTo WriteLoggedRemoveChar_Err
        '</EhHeader>

        Dim A As Long
    
100     Call Writer.WriteInt(ServerPacketID.LoggedRemoveChar)
102     Call Writer.WriteInt(SlotUserName)
    
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteLoggedRemoveChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteLoggedRemoveChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleLoginChar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLoginChar_Err
        '</EhHeader>


        Dim UserName As String

        Dim Version  As String

        Dim Key      As String
    
        Dim Slot As Byte
    
100     Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
102     UserName = Reader.ReadString8()
104     Key = Reader.ReadString8
106     Slot = Reader.ReadInt8
    
108     If Not VersionOK(Version) Then
110         Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
112     ElseIf PuedeConectarPersonajes = 0 Then
114         Call Protocol.Kick(UserIndex, "No está permitido el ingreso de personajes al juego.")
116     ElseIf CheckUserLogged(UCase$(UserName)) Then
118         Call WriteErrorMsg(UserIndex, "El personaje se encuentra online.")
        Else
120         Call mAccount.LoginAccount_Char(UserIndex, UserName, Key, Slot, False)
        End If


        '<EhFooter>
        Exit Sub

HandleLoginChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLoginChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleDisconnectForced(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleDisconnectForced_Err
        '</EhHeader>

        Dim Account   As String
        Dim Key       As String
        Dim Version   As String

100     Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
102     Account = Reader.ReadString8()
104     Key = Reader.ReadString8()
    
106     If Not VersionOK(Version) Then
108         Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
        Else
110         Call mAccount.DisconnectForced(UserIndex, LCase$(Account), Key)
        End If


        '<EhFooter>
        Exit Sub

HandleDisconnectForced_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleDisconnectForced " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub HandleLoginCharNew(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLoginCharNew_Err
        '</EhHeader>


        Dim Key       As String

        Dim Version   As String

        Dim UserName  As String

        Dim UserClase As Byte

        Dim UserRaza  As Byte

        Dim UserSexo  As Byte
    
        Dim UserHead As Integer
    
        Dim Slot As Byte
    
100     Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
102     UserName = Reader.ReadString8()
104     UserClase = Reader.ReadInt8()
106     UserRaza = Reader.ReadInt8()
108     UserSexo = Reader.ReadInt8()
110     UserHead = Reader.ReadInt16()
    
    
114     If Not VersionOK(Version) Then
116         Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
        Else
118         Call mAccount.LoginAccount_CharNew(UserIndex, UserName, UserClase, UserRaza, UserSexo, UserHead)
        End If


        '<EhFooter>
        Exit Sub

HandleLoginCharNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLoginCharNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleLoginName(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLoginName_Err
        '</EhHeader>

        Dim Version   As String

        Dim UserName  As String
    
100     Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
102     UserName = Reader.ReadString8()
    
        #If Classic = 1 Then
        Exit Sub
    #End If
    
104     If Not VersionOK(Version) Then
106         Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
        Else
108         Call mAccount.LoginAccount_ChangeAlias(UserIndex, UserName)

        End If
        '<EhFooter>
        Exit Sub

HandleLoginName_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLoginName " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleLoginRemove(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleLoginRemove_Err
        '</EhHeader>

        Dim Key      As String

        Dim Version  As String
    
        Dim Slot As Byte
    
100     Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
102     Key = Reader.ReadString8
104     Slot = Reader.ReadInt8
    
106     If Not VersionOK(Version) Then
108         Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
        Else
110         Call mAccount.LoginAccount_Remove(UserIndex, Key, Slot)
        End If


        '<EhFooter>
        Exit Sub

HandleLoginRemove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleLoginRemove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



Public Sub HandleMercader_New(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleMercader_New_Err
        '</EhHeader>

        Dim Key      As String, Passwd As String
    
        Dim Chars()  As Byte
    
        Dim Mercader As tMercaderChar
    
        Dim Gld      As Long, Dsp As Long
    
        Dim A        As Long, Desc As String
    
        Dim SaleCost As Long
    
        Dim Blocked As Byte
    
        Dim CantChars As Byte
        
        Dim SlotMercader As Integer 'Slot al que queremos ofrecer

102     Passwd = Reader.ReadString8
104     Key = Reader.ReadString8
          
        SlotMercader = Reader.ReadInt16
          
106     Gld = Reader.ReadInt32
        Dsp = Reader.ReadInt32
        Desc = Reader.ReadString8
        
108     SaleCost = Reader.ReadInt32
110     Blocked = Reader.ReadInt8
    
112     Call Reader.ReadSafeArrayInt8(Chars)
            
114     If Not StrComp(Key, UserList(UserIndex).Account.Key) = 0 Then
116         Call WriteErrorMsg(UserIndex, "Has escrito una clave de seguridad erronea.")
            Exit Sub
        End If
        
        If Not StrComp(Passwd, UserList(UserIndex).Account.Passwd) = 0 Then
            Call WriteErrorMsg(UserIndex, "Has escrito una contraseña incorrecta.")
            Exit Sub
        End If
        
        If SlotMercader < 0 Or SlotMercader > mMao.MERCADER_MAX_LIST Then Exit Sub
        
118     If MercaderActivate Then
        
120         Mercader.Account = UserList(UserIndex).Account.Email
122         Mercader.Gld = Gld
            Mercader.Dsp = Dsp
            Mercader.Desc = Desc
                
             If SlotMercader = 0 Then
                 If UserList(UserIndex).Account.MercaderSlot > 0 Then
                        Call WriteErrorMsg(UserIndex, "¡Ya tienes una publicación vigente! Elimina la que tienes para crear otra...")
                        Exit Sub
                End If
            Else
                If StrComp(MercaderList(SlotMercader).Chars.Account, UserList(UserIndex).Account.Email) = 0 Then
                    Call WriteErrorMsg(UserIndex, "¡No puedes hacer intercambios contigo mismo!")
                    Exit Sub
                End If
            End If
                
            If SlotMercader = 0 Then
136             Call mMao.Mercader_AddList(UserIndex, Chars, Mercader, Blocked)
            Else
                Call mMao.Mercader_AddOffer(UserIndex, Chars, SlotMercader, Mercader, Blocked)

            End If
        Else
138         Call WriteErrorMsg(UserIndex, "Mercado desactivado temporalmente.")
        End If

        '<EhFooter>
        Exit Sub

HandleMercader_New_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleMercader_New " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleMercader_Required(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo HandleMercader_Required_Err

        '</EhHeader>
    
        Dim Version  As String

        Dim Required As Byte

        Dim Value    As Long, Value1 As Long

100     Required = Reader.ReadInt
102     Value = Reader.ReadInt
104     Value1 = Reader.ReadInt

106     Select Case Required

            Case 0 ' Remover publicación.

                If UserList(UserIndex).Account.MercaderSlot > 0 Then
                    Call Mercader_Remove(UserList(UserIndex).Account.MercaderSlot, UserList(UserIndex).Account.Email)
                    Call WriteErrorMsg(UserIndex, "¡Has eliminado la publicación!")
                    Call WriteAccountInfo(UserIndex)

                End If
                
108         Case 1 ' Enviar Lista del Mercado.

110             If Value <= 0 Then Exit Sub
112             If Value1 <= 0 Then Exit Sub
114             If Value > MERCADER_MAX_LIST Then Value = MERCADER_MAX_LIST
116             If Value1 > MERCADER_MAX_LIST Then Value1 = MERCADER_MAX_LIST
118             Call WriteMercader_List(UserIndex, Value, Value1, 0)

120         Case 2 ' Enviar Información del listado seleccionado.

122             If Value <= 0 Or Value > MERCADER_MAX_LIST Then Exit Sub
124             If Value1 <= 0 Or Value1 > ACCOUNT_MAX_CHARS Then Exit Sub
                  If MercaderList(Value).Chars.Account = vbNullString Then Exit Sub
                  If MercaderList(Value).Chars.Count = 0 Then Exit Sub
126             Call WriteMercader_ListChar(UserIndex, Value, Value1, False)

128         Case 3 ' Enviar la lista de ofertas
                If Value <= 0 Or Value > MERCADER_MAX_LIST Then Exit Sub
                    Call WriteMercader_List(UserIndex, 1, 50, Value)
                    
130         Case 4 ' Envia la información de las ofertas

                If Value <= 0 Or Value > MERCADER_MAX_LIST Then Exit Sub
                If Value1 <= 0 Or Value1 > MERCADER_MAX_OFFER Then Exit Sub
                If MercaderList(Value).Offer(Value1).Account = vbNullString Then Exit Sub
                If MercaderList(Value).Offer(Value1).Count = 0 Then Exit Sub
                Call WriteMercader_ListChar(UserIndex, Value, Value1, True)
            
132         Case 5 ' Acepta una oferta recibida
                  If Value <= 0 Or Value > MERCADER_MAX_OFFER Then Exit Sub
                  If UserList(UserIndex).Account.MercaderSlot = 0 Then Exit Sub
                  If MercaderList(UserList(UserIndex).Account.MercaderSlot).Offer(Value).Account = vbNullString Then Exit Sub
                 Call mMao.Mercader_AcceptOffer(UserIndex, UserList(UserIndex).Account.MercaderSlot, Value)
        End Select

        '<EhFooter>
        Exit Sub

HandleMercader_Required_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMercader_Required " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub WriteMercader_List(ByVal UserIndex As Integer, _
                              ByVal aBound As Integer, _
                              ByVal bBound As Integer, _
                              ByVal MercaderSlot As Integer)
        '<EhHeader>
        On Error GoTo WriteMercader_List_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.Mercader_List)
    
        Dim A    As Long, B As Long
        Dim Text As String
    
        Dim Mercader As tMercaderChar
    
102     Call Writer.WriteInt16(aBound)
104     Call Writer.WriteInt16(bBound)

        Call Writer.WriteInt16(UserList(UserIndex).Account.MercaderSlot)
    
106     For A = aBound To bBound
108         If MercaderSlot = 0 Then
110             Mercader = MercaderList(A).Chars
            Else
112             Mercader = MercaderList(MercaderSlot).Offer(A)
            End If
        
114         With Mercader
116             Call Writer.WriteInt16(A)
118             Call Writer.WriteInt8(.Count)
                Call Writer.WriteString8(.Desc)
                Call Writer.WriteInt32(.Dsp)
120             Call Writer.WriteInt32(.Gld)
            
122             For B = 1 To .Count
124                 Call Writer.WriteString8(.NameU(B))
126                 Call Writer.WriteInt8(.Info(B).Class)
128                 Call Writer.WriteInt8(.Info(B).Raze)
                
130                 Call Writer.WriteInt8(.Info(B).Elv)
132                 Call Writer.WriteInt32(.Info(B).Exp)
134                 Call Writer.WriteInt32(.Info(B).Elu)
                    
136                 Call Writer.WriteInt16(.Info(B).Hp)
138                 Call Writer.WriteInt8(.Info(B).Constitucion)

                    Call Writer.WriteInt16(.Info(B).Body)
                    Call Writer.WriteInt16(.Info(B).Head)
                    Call Writer.WriteInt16(.Info(B).Weapon)
                    Call Writer.WriteInt16(.Info(B).Shield)
                    Call Writer.WriteInt16(.Info(B).Helm)

140             Next B

            End With
142     Next A
    
144     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteMercader_List_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMercader_List " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteMercader_ListChar(ByVal UserIndex As Integer, _
                                  ByVal Slot As Integer, _
                                  ByVal SlotChar As Integer, _
                                  ByVal InfoOffer As Boolean)
        '<EhHeader>
        On Error GoTo WriteMercader_ListChar_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.Mercader_ListInfo)
    
        Dim A        As Long, B As Long

        Dim Text     As String

        Dim Mercader As tMercaderChar
    
102     Call Writer.WriteInt16(SlotChar)
    
104     If Not InfoOffer Then
106         Mercader = MercaderList(Slot).Chars
        Else
108         Mercader = MercaderList(Slot).Offer(SlotChar)

        End If
        
110     With Mercader
112         Call Writer.WriteString8(.NameU(SlotChar))
114         Call Writer.WriteString8(GuildsInfo(.Info(SlotChar).GuildIndex).Name)
        
116         Call Writer.WriteInt32(.Info(SlotChar).Gld)
        
118         Call Writer.WriteInt16(.Info(SlotChar).Body)
120         Call Writer.WriteInt16(.Info(SlotChar).Head)
122         Call Writer.WriteInt16(.Info(SlotChar).Weapon)
124         Call Writer.WriteInt16(.Info(SlotChar).Shield)
126         Call Writer.WriteInt16(.Info(SlotChar).Helm)
        
128         Call Writer.WriteInt8(.Info(SlotChar).Faction)
130         Call Writer.WriteInt8(.Info(SlotChar).FactionRange)
132         Call Writer.WriteInt16(.Info(SlotChar).FragsCiu)
134         Call Writer.WriteInt16(.Info(SlotChar).FragsCri)
        
136         For A = 1 To MAX_INVENTORY_SLOTS
138             Call Writer.WriteInt16(.Info(SlotChar).Object(A).ObjIndex)
140             Call Writer.WriteInt16(.Info(SlotChar).Object(A).Amount)
142         Next A
            
144         For A = 1 To MAX_BANCOINVENTORY_SLOTS
146             Call Writer.WriteInt16(.Info(SlotChar).Bank(A).ObjIndex)
148             Call Writer.WriteInt16(.Info(SlotChar).Bank(A).Amount)
150         Next A
        
152         For A = 1 To 35

154             If .Info(SlotChar).Spells(A) > 0 Then
156                 Call Writer.WriteString8(Hechizos(.Info(SlotChar).Spells(A)).Nombre)
                Else
158                 Call Writer.WriteString8(vbNullString)

                End If

160         Next A
        
162         For A = 1 To NUMSKILLS
164             Call Writer.WriteInt8(.Info(SlotChar).Skills(A))
166         Next A
        
        End With
    
168     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteMercader_ListChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMercader_ListChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteMercader_ListOffer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteMercader_ListOffer_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.Mercader_ListOffer)
    

    
102     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteMercader_ListOffer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMercader_ListOffer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



Private Sub HandleForgive_Faction(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleForgive_Faction_Err
        '</EhHeader>

100     With UserList(UserIndex)
        
            'Validate target NPC
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

106         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

108         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
110             Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

112         If .Faction.Status > 0 Then
114             Call WriteConsoleMsg(UserIndex, "Ya eres miembro de una facción y espero que sea la nuestra, sino mis guardias te atacaran!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
116         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
118             If Escriminal(UserIndex) Then
120                 Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                    Exit Sub

                End If
            
122             If Not TieneObjetos(1086, 1, UserIndex) Then
124                 Call WriteConsoleMsg(UserIndex, "¡No has reclamado tu recompensa! Debes hacer la misión que te otorga el fragmento necesario para el perdón", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub
                End If
            
126             Call QuitarObjetos(1086, 1, UserIndex)
            
128             UserList(UserIndex).Faction.FragsCiu = 0
            
            Else

130             If Not Escriminal(UserIndex) Then
132                 Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                    Exit Sub

                End If
            
134             If Not TieneObjetos(1087, 1, UserIndex) Then
136                 Call WriteConsoleMsg(UserIndex, "¡No has reclamado tu recompensa! Debes hacer la misión que te otorga el fragmento necesario para el perdón", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub
                End If
            
138             Call QuitarObjetos(1087, 1, UserIndex)
            End If
        
        
140         Call Faction_RemoveUser(UserIndex)
142         Call WriteConsoleMsg(UserIndex, "¡Te hemos perdonado, pero no abuses de nuestra bondad. Nuestras tropas son fieles y no toleran estupideces!", FontTypeNames.FONTTYPE_DEMONIO)
        
        End With

        '<EhFooter>
        Exit Sub

HandleForgive_Faction_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleForgive_Faction " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleMap_RequiredInfo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleMap_RequiredInfo_Err
        '</EhHeader>
    
        Dim Map As Integer
    
100     Map = Reader.ReadInt
    
102     If Map = 0 Or Map > NumMaps Then Exit Sub
    
104     Call WriteMiniMap_InfoCriature(UserIndex, Map)

        '<EhFooter>
        Exit Sub

HandleMap_RequiredInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleMap_RequiredInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteRender_CountDown(ByVal UserIndex As Integer, ByVal CountDown As Long)
        '<EhHeader>
        On Error GoTo WriteRender_CountDown_Err
        '</EhHeader>
100     Call SendData(ToOne, UserIndex, PrepareMessageRender_CountDown(CountDown))
        '<EhFooter>
        Exit Sub

WriteRender_CountDown_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteRender_CountDown " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function PrepareMessageRender_CountDown(ByVal Time As Long) As String
        '<EhHeader>
        On Error GoTo PrepareMessageRender_CountDown_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.Render_CountDown)
102     Call Writer.WriteInt(Time)
        
        '<EhFooter>
        Exit Function

PrepareMessageRender_CountDown_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageRender_CountDown " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Private Sub WriteMiniMap_InfoCriature(ByVal UserIndex As Integer, _
                                     ByVal Map As Integer)
        '<EhHeader>
        On Error GoTo WriteMiniMap_InfoCriature_Err
        '</EhHeader>

        Dim A    As Long, B As Long

        Dim Str  As String

        Dim Temp As String

100     Call Writer.WriteInt(ServerPacketID.MiniMap_InfoCriature)
102     Call Writer.WriteInt(Map)
104     Call Writer.WriteInt(MiniMap(Map).NpcsNum)
106     Call Writer.WriteString8(MiniMap(Map).Name)
108     Call Writer.WriteBool(MiniMap(Map).Pk)
110     Call Writer.WriteInt(MiniMap(Map).LvlMin)
112     Call Writer.WriteInt(MiniMap(Map).LvlMax)
            
114     If MiniMap(Map).NpcsNum Then
116         With MiniMap(Map)
        
118             For A = 1 To MiniMap(Map).NpcsNum
120                 Call Writer.WriteInt16(MiniMap(Map).Npcs(A).NpcIndex)
122                 Call Writer.WriteString8(MiniMap(Map).Npcs(A).Name)
124                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).Body)
126                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).Head)
128                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).Hp)
130                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).MinHit)
132                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).MaxHit)
134                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).Exp)
136                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).Gld)
138                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).Eldhir)
                
                    'Spells
140                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).NroSpells)
                
142                 If MiniMap(Map).Npcs(A).NroSpells Then
    
144                     For B = 1 To MiniMap(Map).Npcs(A).NroSpells
146                         Call Writer.WriteString8(Hechizos(MiniMap(Map).Npcs(A).Spells(B)).Nombre)
148                     Next B
    
                    End If
                
                    ' Inventario de la Criatura
150                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).NroItems)
                
152                 For B = 1 To MiniMap(Map).Npcs(A).NroItems
154                     Temp = vbNullString
                        
156                     If MiniMap(Map).Npcs(A).Invent.Object(B).ObjIndex > 0 Then
158                         Temp = ObjData(MiniMap(Map).Npcs(A).Invent.Object(B).ObjIndex).Name
                        End If
                        
160                     Call Writer.WriteString8(Temp)
162                     Call Writer.WriteInt(MiniMap(Map).Npcs(A).Invent.Object(B).Amount)
164                 Next B
                
                    ' Drops de la Criatura
166                 Call Writer.WriteInt(MiniMap(Map).Npcs(A).NroDrops)
                
168                 For B = 1 To MiniMap(Map).Npcs(A).NroDrops
170                     Temp = vbNullString
                        
172                     If MiniMap(Map).Npcs(A).Drop(B).ObjIndex > 0 Then
174                         Temp = ObjData(MiniMap(Map).Npcs(A).Drop(B).ObjIndex).Name
                        End If
                        
176                     Call Writer.WriteString8(Temp)
178                     Call Writer.WriteInt(MiniMap(Map).Npcs(A).Drop(B).Amount)
180                     Call Writer.WriteInt(MiniMap(Map).Npcs(A).Drop(B).Probability)
182                 Next B
            
184             Next A
        
            End With
        
        End If
        
186     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteMiniMap_InfoCriature_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteMiniMap_InfoCriature " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleWherePower(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleWherePower_Err
        '</EhHeader>
    
100     If Power.UserIndex = 0 Then
102         Call WriteConsoleMsg(UserIndex, "Ningún usuario posee el don.", FontTypeNames.FONTTYPE_INFORED)
        Else
104         Call WriteConsoleMsg(UserIndex, "El poseedor del poder es el personaje " & UserList(Power.UserIndex).Name & _
                    " en el mapa " & MapInfo(UserList(Power.UserIndex).Pos.Map).Name, FontTypeNames.FONTTYPE_INFOGREEN)
        End If
        '<EhFooter>
        Exit Sub

HandleWherePower_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleWherePower " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleAuction_New(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAuction_New_Err
        '</EhHeader>

        Dim Slot As Byte
        Dim Amount As Integer
        Dim Gld As Long
        Dim Eldhir As Long
    
100     Slot = Reader.ReadInt
102     Amount = Reader.ReadInt
104     Gld = Reader.ReadInt
106     Eldhir = Reader.ReadInt
    
108     If Slot <= 0 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
110     If Amount <= 0 Or Amount > 10000 Then Exit Sub
112     If Gld < 0 Or Gld > 100000000 Then Exit Sub
114     If Eldhir < 0 Or Eldhir > 1000 Then Exit Sub
116     If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
118     If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then Exit Sub
120     If UserList(UserIndex).flags.Bronce = 0 Then
122         Call WriteConsoleMsg(UserIndex, "Debes ser [BRONCE] para poder subastar objetos.", FontTypeNames.FONTTYPE_USERBRONCE)
            Exit Sub
        End If
    
124     Call Auction_CreateNew(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex, Amount, Gld, Eldhir)
        '<EhFooter>
        Exit Sub

HandleAuction_New_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAuction_New " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleAuction_Info(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAuction_Info_Err
        '</EhHeader>
    
100     If Auction.ObjIndex = 0 Then
102         Call WriteConsoleMsg(UserIndex, "¡No hay ninguna subasta en trámite!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
104     Call WriteConsoleMsg(UserIndex, "El personaje " & Auction.Name & " está subastando " & ObjData(Auction.ObjIndex).Name & " (x" & Auction.Amount & "). Deberás ofrecer como mínimo: " & Auction.Offer.Gld * 1.1 & " Monedas de Oro Y " & Auction.Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_INFOGREEN)
        '<EhFooter>
        Exit Sub

HandleAuction_Info_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAuction_Info " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleAuction_Offer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAuction_Offer_Err
        '</EhHeader>
    
        Dim Gld As Long
        Dim Eldhir As Long
    
100     Gld = Reader.ReadInt
102     Eldhir = Reader.ReadInt
    
104     If Gld < 0 Or Gld > 1000000000 Then Exit Sub
106     If Eldhir < 0 Or Eldhir > 5000 Then Exit Sub
    
108     Call mAuction.Auction_Offer(UserIndex, Gld, Eldhir)
        '<EhFooter>
        Exit Sub

HandleAuction_Offer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAuction_Offer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Private Sub HandleGoInvation(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleGoInvation_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            Dim Slot As Byte
        
102         Slot = Reader.ReadInt8
        
104         If Slot <= 0 Or Slot > UBound(Invations) Then Exit Sub
        
106         If Not .Pos.Map = Ullathorpe.Map Then
108             Call WriteConsoleMsg(UserIndex, "Solo puedes ingresar a la invasión estando en Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
110         If Invations(Slot).Run = False Then
112             Call WriteConsoleMsg(UserIndex, "El evento no se encuentra disponible.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
114         Call EventWarpUser(UserIndex, Invations(Slot).InitialMap, Invations(Slot).InitialX, Invations(Slot).InitialY)
116         Call WriteConsoleMsg(UserIndex, "¡Bienvenido a " & Invations(Slot).Name & "! Esperemos que te diviertas y compartas tu experiencia con el resto de los usuarios. ¡Suerte!", FontTypeNames.FONTTYPE_INVASION)
        
118         .Counters.Shield = 3
        
        End With
    
        '<EhFooter>
        Exit Sub

HandleGoInvation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleGoInvation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleSendDataUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSendDataUser_Err
        '</EhHeader>
    
        Dim UserName As String
        Dim tUser As Integer
        Dim Message As String
    
100     UserName = Reader.ReadString8
    
102     If Not EsGmPriv(UserIndex) Then Exit Sub
    
104     tUser = NameIndex(UserName)
    
106     If tUser > 0 Then
108         With UserList(tUser)
110             Message = "'DATOS DE " & UserName & "'"
112             Message = Message & vbCrLf & "IP PUBLICA: " & .Account.Sec.IP_Public
114             Message = Message & vbCrLf & "IP ADDRESS: " & .Account.Sec.IP_Address
116             Message = Message & vbCrLf & "IP LOCAL: " & .Account.Sec.IP_Local
118             Message = Message & vbCrLf & "MAC ADDRESS: " & .Account.Sec.SERIAL_MAC
120             Message = Message & vbCrLf & "DISCO: " & .Account.Sec.SERIAL_DISK
122             Message = Message & vbCrLf & "BIOS: " & .Account.Sec.SERIAL_BIOS
124             Message = Message & vbCrLf & "MOTHERBOARD: " & .Account.Sec.SERIAL_MOTHERBOARD
126             Message = Message & vbCrLf & "PROCESSOR: " & .Account.Sec.SERIAL_PROCESSOR
128             Message = Message & vbCrLf & "SYSTEM DATA " & .Account.Sec.SYSTEM_DATA
            End With
        
130         Call WriteConsoleMsg(UserIndex, Message, FontTypeNames.FONTTYPE_INFOGREEN)
        Else
132         Call WriteConsoleMsg(UserIndex, "El personaje está offline.", FontTypeNames.FONTTYPE_INFORED)
        End If
    
        '<EhFooter>
        Exit Sub

HandleSendDataUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSendDataUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleSearchDataUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleSearchDataUser_Err
        '</EhHeader>
        Dim Data As String
        Dim Selected As eSearchData
    
100     Selected = Reader.ReadInt8
102     Data = Reader.ReadString8
    
104     If Not EsGmPriv(UserIndex) Then Exit Sub
106     If Data = vbNullString Then Exit Sub
108     If Selected <= 0 Or Selected > 3 Then Exit Sub
    
110     Call Security_SearchData(UserIndex, Selected, Data)
        '<EhFooter>
        Exit Sub

HandleSearchDataUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleSearchDataUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleChangeModoArgentum(ByVal UserIndex As Integer)
    If Not EsGmPriv(UserIndex) Then Exit Sub
    
    ' Cambia el Uso de Paquetes
    If PacketUseItem = ClientPacketID.UseItem Then
        PacketUseItem = ClientPacketID.UseItemTwo
        
        EsModoEvento = 1
        Call WriteConsoleMsg(UserIndex, "¡Has pasado al MODO EVENTO!", FontTypeNames.FONTTYPE_INFO)
    Else
        PacketUseItem = ClientPacketID.UseItem
        
        Call WriteConsoleMsg(UserIndex, "¡Has vuelto al MODO Default!", FontTypeNames.FONTTYPE_INFO)
        EsModoEvento = 0
        
    End If
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageUpdateEvento(EsModoEvento))
End Sub

Public Function WriteUpdateEffect(ByVal UserIndex As Integer) As String
        '<EhHeader>
        On Error GoTo WriteUpdateEffect_Err
        '</EhHeader>

        '***************************************************
        ' Actualiza distintos tipos de efectos
        ' Efecto n°1: Veneno (Efecto Verde)
        '
        '***************************************************

100     Call Writer.WriteInt(ServerPacketID.UpdateEffectPoison)
        
102     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Function

WriteUpdateEffect_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateEffect " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Prepares the "CreateFXMap" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFXMap(ByVal X As Byte, _
                                          ByVal Y As Byte, _
                                          ByVal FX As Integer, _
                                          ByVal FXLoops As Integer) As String
        '<EhHeader>
        On Error GoTo PrepareMessageCreateFXMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification:
        'Prepares the "CreateFXMap" message and returns it
        '***************************************************
    
100     Call Writer.WriteInt(ServerPacketID.CreateFXMap)
102     Call Writer.WriteInt(X)
104     Call Writer.WriteInt(Y)
106     Call Writer.WriteInt(FX)
108     Call Writer.WriteInt(FXLoops)

        '<EhFooter>
        Exit Function

PrepareMessageCreateFXMap_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageCreateFXMap " & _
               "at line " & Erl
        
        '</EhFooter>
End Function



Private Sub HandleEvents_DonateObject(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleEvents_DonateObject_Err
        '</EhHeader>

        Dim Slot As Byte
        Dim SlotEvent As Byte
        Dim Amount As Integer

100     SlotEvent = Reader.ReadInt8
102     Slot = Reader.ReadInt8
104     Amount = Reader.ReadInt16
    
106     If SlotEvent <= 0 Or SlotEvent > UBound(Events) Then Exit Sub
108     If Slot <= 0 Or Slot >= MAX_INVENTORY_SLOTS Then Exit Sub
110     If Amount > UserList(UserIndex).Invent.Object(Slot).Amount Or Amount >= MAX_INVENTORY_OBJS Then Exit Sub
112     If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then Exit Sub
114     If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Donable = 0 Then
116         Call WriteConsoleMsg(UserIndex, "¡Ha ha ha tu objeto no es tolerado aquí!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
118     Call EventosDS_Reward.Events_Reward_Add(UserIndex, SlotEvent, Slot, Amount)
   
        '<EhFooter>
        Exit Sub

HandleEvents_DonateObject_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleEvents_DonateObject " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Function PrepareMessageRenderConsole(ByVal Text As String, _
                                            ByVal DamageType As eDamageType, _
                                            ByVal Duration As Long, _
                                            ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageRenderConsole_Err
        '</EhHeader>
 
100     Writer.WriteInt ServerPacketID.RenderConsole
102     Writer.WriteString8 Text
104     Writer.WriteInt8 DamageType
106     Writer.WriteInt32 Duration
108     Writer.WriteInt8 Slot

        '<EhFooter>
        Exit Function

PrepareMessageRenderConsole_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageRenderConsole " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub WriteViewListQuest(ByVal UserIndex As Integer, _
                                ByRef Quest() As Byte, _
                                ByVal NameNpc As String)
        '<EhHeader>
        On Error GoTo WriteViewListQuest_Err
        '</EhHeader>
    
        Dim A As Long, B As Long
    
    
100     Call Writer.WriteInt(ServerPacketID.ViewListQuest)
102     Call Writer.WriteInt8(UBound(Quest))
104     Call Writer.WriteString8(NameNpc)
    
106     For A = LBound(Quest) To UBound(Quest)
108         Call Writer.WriteInt8(Quest(A))
110     Next A

112     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteViewListQuest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteViewListQuest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdateUserDead(ByVal UserIndex As Integer, _
                                ByVal UserMuerto As Byte)
        '<EhHeader>
        On Error GoTo WriteUpdateUserDead_Err
        '</EhHeader>
                                
100     Call Writer.WriteInt(ServerPacketID.UpdateUserDead)
102     Call Writer.WriteInt8(UserMuerto)
    
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUpdateUserDead_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateUserDead " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteQuestInfo(ByVal UserIndex As Integer, _
                          ByVal Visible As Boolean, _
                          ByVal Slot As Integer)
        '<EhHeader>
        On Error GoTo WriteQuestInfo_Err
        '</EhHeader>
                                
100     Call Writer.WriteInt(ServerPacketID.QuestInfo)
        
        Dim A          As Long
        
        Dim i          As Long

        Dim QuestIndex As Integer
    
102     Call Writer.WriteBool(Visible)
104     Call Writer.WriteInt16(Slot)
    
106     If Slot <> 0 Then
108         QuestIndex = UserList(UserIndex).QuestStats(Slot).QuestIndex
                
110         Call Writer.WriteInt16(QuestIndex)
                
112         If UserList(UserIndex).QuestStats(Slot).QuestIndex > 0 Then
            
114             If QuestList(QuestIndex).RequiredNPCs > 0 Then
    
116                 For i = LBound(UserList(UserIndex).QuestStats(Slot).NPCsKilled) To UBound(UserList(UserIndex).QuestStats(Slot).NPCsKilled)
118                     Call Writer.WriteInt32(UserList(UserIndex).QuestStats(Slot).NPCsKilled(i))
120                 Next i
    
                End If
                        
122             If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then
    
124                 For i = LBound(UserList(UserIndex).QuestStats(Slot).ObjsSale) To UBound(UserList(UserIndex).QuestStats(Slot).ObjsSale)
126                     Call Writer.WriteInt32(UserList(UserIndex).QuestStats(Slot).ObjsSale(i))
128                 Next i
    
                End If
                        
130             If QuestList(QuestIndex).RequiredChestOBJs > 0 Then
    
132                 For i = LBound(UserList(UserIndex).QuestStats(Slot).ObjsPick) To UBound(UserList(UserIndex).QuestStats(Slot).ObjsPick)
134                     Call Writer.WriteInt32(UserList(UserIndex).QuestStats(Slot).ObjsPick(i))
136                 Next i
    
                End If
                        
            End If

        Else
    
138         For A = 1 To MAXUSERQUESTS
140             QuestIndex = UserList(UserIndex).QuestStats(A).QuestIndex
                
142             Call Writer.WriteInt16(QuestIndex)
                
144             If UserList(UserIndex).QuestStats(A).QuestIndex > 0 Then
            
146                 If QuestList(QuestIndex).RequiredNPCs > 0 Then
    
148                     For i = LBound(UserList(UserIndex).QuestStats(A).NPCsKilled) To UBound(UserList(UserIndex).QuestStats(A).NPCsKilled)
150                         Call Writer.WriteInt32(UserList(UserIndex).QuestStats(A).NPCsKilled(i))
152                     Next i
    
                    End If
                        
154                 If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then
    
156                     For i = LBound(UserList(UserIndex).QuestStats(A).ObjsSale) To UBound(UserList(UserIndex).QuestStats(A).ObjsSale)
158                         Call Writer.WriteInt32(UserList(UserIndex).QuestStats(A).ObjsSale(i))
160                     Next i
    
                    End If
                        
162                 If QuestList(QuestIndex).RequiredChestOBJs > 0 Then
    
164                     For i = LBound(UserList(UserIndex).QuestStats(A).ObjsPick) To UBound(UserList(UserIndex).QuestStats(A).ObjsPick)
166                         Call Writer.WriteInt32(UserList(UserIndex).QuestStats(A).ObjsPick(i))
168                     Next i
    
                    End If
                        
                End If

170         Next A
    
        End If

172     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteQuestInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteQuestInfo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub HandleQuestRequired(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleQuestRequired_Err
        '</EhHeader>
    
        Dim Tipo As Byte
100     Tipo = Reader.ReadInt8
    
112     Call WriteQuestInfo(UserIndex, True, 0)
        '<EhFooter>
        Exit Sub

HandleQuestRequired_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleQuestRequired " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdateGlobalCounter(ByVal UserIndex As Integer, ByVal Tipo As Byte, ByVal Counter As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateGlobalCounter_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateGlobalCounter)
    
102     Call Writer.WriteInt8(Tipo)
104     Call Writer.WriteInt16(Counter)

106      Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUpdateGlobalCounter_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateGlobalCounter " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub WriteSendIntervals(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSendIntervals_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.SendIntervals)

102     Call Writer.WriteInt16(IntervaloUserPuedeAtacar)
104     Call Writer.WriteInt16(IntervaloUserPuedeUsar)
106     Call Writer.WriteInt16(IntervaloUserPuedeUsarClick)
108     Call Writer.WriteInt16(2000) ' Actualizar POS
110     Call Writer.WriteInt16(IntervaloUserPuedeCastear)
          Call Writer.WriteInt16(IntervaloUserPuedeShiftear)
112     Call Writer.WriteInt16(IntervaloFlechasCazadores)
114     Call Writer.WriteInt16(IntervaloMagiaGolpe)
116     Call Writer.WriteInt16(IntervaloGolpeMagia)
118     Call Writer.WriteInt16(IntervaloGolpeUsar)
120     Call Writer.WriteInt16(IntervaloUserPuedeTrabajar)
122     Call Writer.WriteInt16(IntervalDrop)
          Call Writer.WriteReal32(IntervaloCaminar)
          
124     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteSendIntervals_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteSendIntervals " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteSendInfoNpc(ByVal UserIndex As Integer, ByVal number As Integer)
        '<EhHeader>
        On Error GoTo WriteSendInfoNpc_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.SendInfoNpc)

102     Call Writer.WriteInt16(number)
   
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteSendInfoNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteSendInfoNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdatePosGuild(ByVal UserIndex As Integer, ByVal SlotMember As Byte, ByVal tUser As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdatePosGuild_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdatePosGuild)

102     Call Writer.WriteInt8(SlotMember)
    
104     If tUser > 0 Then
106         Call Writer.WriteInt8(UserList(tUser).Pos.X)
108         Call Writer.WriteInt8(UserList(tUser).Pos.Y)
        Else
110         Call Writer.WriteInt8(0)
112         Call Writer.WriteInt8(0)
        End If
    
114     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUpdatePosGuild_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdatePosGuild " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function PrepareUpdateLevelGuild(ByVal LevelGuild As Byte)
        '<EhHeader>
        On Error GoTo PrepareUpdateLevelGuild_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateLevelGuild)
102     Call Writer.WriteInt8(LevelGuild)
        '<EhFooter>
        Exit Function

PrepareUpdateLevelGuild_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareUpdateLevelGuild " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub WriteUpdateStatusMAO(ByVal UserIndex As Integer, ByVal Status As Byte)
        '<EhHeader>
        On Error GoTo WriteUpdateStatusMAO_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateStatusMAO)
102     Call Writer.WriteInt8(Status)
104     Call SendData(ToOne, UserIndex, vbNullString)

        '<EhFooter>
        Exit Sub

WriteUpdateStatusMAO_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateStatusMAO " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub HandleChangeClass(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeClass_Err
        '</EhHeader>
        Dim Clase As Byte
        Dim Raza As Byte
        Dim Genero As Byte
    
100     Clase = Reader.ReadInt8
102     Raza = Reader.ReadInt8
104     Genero = Reader.ReadInt8
    
    
106     If Clase <= 0 Or Clase > NUMCLASES Then Exit Sub
108     If Raza <= 0 Or Raza > NUMRAZAS Then Exit Sub
110     If Genero <= 0 Or Genero > 2 Then Exit Sub
    
        #If Classic = 1 Then
            Exit Sub
        #End If
    
112     With UserList(UserIndex)
114         .Clase = Clase
116         .Raza = Raza
118         .Genero = Genero
120         .flags.Muerto = 0
        
        
        
122         Call InitialUserStats(UserList(UserIndex))
124         Call UserLevelEditation(UserList(UserIndex), STAT_MAXELV, 0)
              Call WriteUpdateUserStats(UserIndex)
              Call WriteUpdateHungerAndThirst(UserIndex)
126         Call LoadSetInitial_Class(UserIndex)
            'Call LoadSetInitial_Class(UserIndex)
128         Call WriteConsoleMsg(UserIndex, "Ahora eres un " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & ".", FontTypeNames.FONTTYPE_INFO)
        End With
    
        '<EhFooter>
        Exit Sub

HandleChangeClass_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeClass " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function PrepareMessageUpdateOnline() As String
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateOnline_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.UpdateOnline)
102     Call Writer.WriteInt16(NumUsers + UsersBot)
        
        '<EhFooter>
        Exit Function

PrepareMessageUpdateOnline_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageUpdateOnline " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' SEGURIDAD
Public Function PrepareMessageUpdateEvento(ByVal ModoEvento As Byte) As String
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateEvento_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.UpdateEvento)
102     Call Writer.WriteInt8(ModoEvento)
        
        '<EhFooter>
        Exit Function

PrepareMessageUpdateEvento_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageUpdateEvento " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Sub HandleModoStreamer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleModoStreamer_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
                If .flags.StreamUrl = vbNullString Then
                    Call WriteConsoleMsg(UserIndex, "Por favor setea primero una URL con el comando /STREAMLINK. ¡Vende tu contenido! Sé hábil para poner alguna frase que haga que las personas ingresen a tu canal haciendo clic! NO muy largo.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub
                End If
            
102             .flags.ModoStream = Not .flags.ModoStream
            
104             If .flags.ModoStream Then
106                 Call WriteMultiMessage(UserIndex, eMessages.ModoStreamOn)
                Else
108                 Call WriteMultiMessage(UserIndex, eMessages.ModoStreamOff)
                End If
                
               ' Call Streamer_Can(UserIndex)
                
            
110             'Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
    
        End With
  
        '<EhFooter>
        Exit Sub

HandleModoStreamer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleModoStreamer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleStreamerSetLink(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleStreamerSetLink_Err
        '</EhHeader>
    
        Dim Url As String
    
100     Url = Reader.ReadString8
    
    
102     With UserList(UserIndex)
104         .flags.StreamUrl = Url
106         Call WriteConsoleMsg(UserIndex, "¡La URL del Twitch ha pasado a ser " & Url & "!", FontTypeNames.FONTTYPE_INFOGREEN)
        
108         Call Logs_Security(eSecurity, eAntiHack, "La cuenta " & .Account.Email & " con personaje: " & .Name & " ha cambiado su link de Twitch a " & Url & ".")
        End With
    
    
   
        '<EhFooter>
        Exit Sub

HandleStreamerSetLink_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleStreamerSetLink " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Private Sub HandleChangeNick(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleChangeNick_Err
        '</EhHeader>
        Dim UserName As String
        Dim Leader As Boolean
    
100     UserName = Reader.ReadString8
102     Leader = Reader.ReadBool
    
104     If Leader Then
106          Call mGuilds.ChangeLeader(UserIndex, UCase$(UserName))
        Else
108          Call mAccount.ChangeNickChar(UserIndex, UCase$(UserName))
        End If
    
   
        '<EhFooter>
        Exit Sub

HandleChangeNick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleChangeNick " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleConfirmTransaccion(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleConfirmTransaccion_Err
        '</EhHeader>
        Dim Waiting As tShopWaiting
    
100     Waiting.Email = Reader.ReadString8
102     Waiting.Promotion = Reader.ReadInt8
104     Waiting.Bank = Reader.ReadString8
    
106     Call mShop.Transaccion_Add(UserIndex, Waiting)
    
        '<EhFooter>
        Exit Sub

HandleConfirmTransaccion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleConfirmTransaccion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleConfirmItem(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleConfirmItem_Err
        '</EhHeader>
        Dim ID As Integer
        Dim PrioriceValue As Byte
        
100     ID = Reader.ReadInt16
          PrioriceValue = Reader.ReadInt8
          
102     If ID <= 0 Or ID > ShopLast Then Exit Sub ' Anti Hack
          If PrioriceValue > 1 Then Exit Sub
          
104     Call mShop.ConfirmItem(UserIndex, ID, PrioriceValue)
    
        '<EhFooter>
        Exit Sub

HandleConfirmItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleConfirmItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleConfirmTier(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleConfirmTier_Err
        '</EhHeader>
        Dim Tier As Byte
    
100     Tier = Reader.ReadInt8
    
102     Call mShop.ConfirmTier(UserIndex, Tier)
        '<EhFooter>
        Exit Sub

HandleConfirmTier_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleConfirmTier " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' SEGURIDAD
Public Function PrepareMessageUpdateMeditation(ByRef MeditationUser() As Integer, ByVal MeditationAnim As Byte) As String
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateMeditation_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.UpdateMeditation)
102     Call Writer.WriteInt16(MeditationAnim)
    
104     Call Writer.WriteInt8(MAX_MEDITATION)
    
        Dim A As Long
    
106     For A = 1 To MAX_MEDITATION
108         Call Writer.WriteInt8(MeditationUser(A))
110     Next A
    
        '<EhFooter>
        Exit Function

PrepareMessageUpdateMeditation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.PrepareMessageUpdateMeditation " & _
               "at line " & Erl
        
        '</EhFooter>
End Function


Private Sub HandleRequiredShopChars(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleRequiredShopChars_Err
        '</EhHeader>
100     If Not Interval_Packet250(UserIndex) Then Exit Sub
102     Call WriteShopChars(UserIndex)
        '<EhFooter>
        Exit Sub

HandleRequiredShopChars_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleRequiredShopChars " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteShopChars(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShopChars_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SendShopChars)
    
        Dim A As Long
    
102     Call Writer.WriteInt8(ShopCharLast)
    
104     For A = 1 To ShopCharLast
106         With ShopChars(A)
108             Call Writer.WriteString8(.Name)
110             Call Writer.WriteInt16(.Dsp)
            
112             Call Writer.WriteInt8(.Elv)
114             Call Writer.WriteInt8(.Porc)
116             Call Writer.WriteInt8(.Class)
118             Call Writer.WriteInt8(.Raze)
120             Call Writer.WriteInt16(.Head)
122             Call Writer.WriteInt16(.Hp)
124             Call Writer.WriteInt16(.Man)
            End With
126     Next A
    
128     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteShopChars_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteShopChars " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleConfirmChar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleConfirmChar_Err
        '</EhHeader>
    
        Dim ID As Byte
    
100     ID = Reader.ReadInt8
    
102     If ID <= 0 Or ID > ShopCharLast Then Exit Sub
    
104     Call mShop.ConfirmChar(UserIndex, ID)
    
        '<EhFooter>
        Exit Sub

HandleConfirmChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleConfirmChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleConfirmQuest(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleConfirmQuest_Err
        '</EhHeader>
    
        Dim Tipo  As Byte

        Dim Quest As Byte
    
100     Tipo = Reader.ReadInt8
102     Quest = Reader.ReadInt8
    
104     Select Case Tipo
    
            Case 1 ' Reclamar Mision
                 
106             If UserList(UserIndex).QuestStats(Quest).QuestIndex Then
108                 If Quests_CheckFinish(UserIndex, Quest) Then
110                     Call mQuests.Quests_Next(UserIndex, Quest)
                    End If
                  
                End If
            
112         Case 2 ' Confirmar para hacer una de las de alto riesgo
    
        End Select
    
        '<EhFooter>
        Exit Sub

HandleConfirmQuest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleConfirmQuest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub WriteUpdateFinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateFinishQuest_Err
        '</EhHeader>
                                              
100     Call Writer.WriteInt(ServerPacketID.UpdateFinishQuest)
102     Call Writer.WriteInt16(QuestIndex)
    
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUpdateFinishQuest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateFinishQuest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub HandleRequiredSkin(ByVal UserIndex As Integer)
        
    If Not Interval_Packet250(UserIndex) Then Exit Sub
        
    Dim ObjIndex As Integer
    Dim Modo As Byte
    
    ObjIndex = Reader.ReadInt16
    Modo = Reader.ReadInt8
    
    If ObjIndex > 0 Then
        If Modo = 3 Then    ' Desequipar
            Call Skins_Desequipar(UserIndex, ObjIndex)
        Else
            Call mSkins.Skins_AddNew(UserIndex, ObjIndex)
        End If
        
    Else
        WriteUpdateDataSkin UserIndex, UserList(UserIndex).Skins.Last
    End If
        
End Sub

Public Sub WriteUpdateDataSkin(ByVal UserIndex As Integer, ByVal Last As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateFinishQuest_Err
        '</EhHeader>
                                              
100       Call Writer.WriteInt(ServerPacketID.UpdateDataSkin)

            Dim Data As tSkins
            Dim A As Long
            
            Data = UserList(UserIndex).Skins
            Call Writer.WriteInt16(Last)
            
            If Last > 0 Then
                Data.Last = Last
                
                If Data.Last > 0 Then
                    
                    For A = 1 To Data.Last
                        Call Writer.WriteInt16(Data.ObjIndex(A))
                    Next A
                End If
            End If
               
            Call Writer.WriteInt16(Data.ArmourIndex)
            Call Writer.WriteInt16(Data.HelmIndex)
            Call Writer.WriteInt16(Data.ShieldIndex)
            Call Writer.WriteInt16(Data.WeaponIndex)
            Call Writer.WriteInt16(Data.WeaponArcoIndex)
            Call Writer.WriteInt16(Data.WeaponDagaIndex)
            
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUpdateFinishQuest_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateFinishQuest " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Mueve al personaje solo desde otro
Public Sub WriteRequiredMoveChar(ByVal UserIndex As Integer, ByVal Heading As Byte)
    
    Call Writer.WriteInt(ServerPacketID.RequiredMoveChar)
    Call Writer.WriteInt8(Heading)
    Call SendData(ToOne, UserIndex, vbNullString)
    
End Sub

Private Sub HandleStreamerBotSetting(ByVal UserIndex As Integer)
    
    Dim Delay As Long ' Delay entre sum & sum
    Dim Mode As eStreamerMode
    Dim DelayIndex As Long
    
    Delay = Reader.ReadInt32
    Mode = Reader.ReadInt8
    DelayIndex = Reader.ReadInt32
    
    If Not EsGm(UserIndex) Then Exit Sub
    
    If Delay < 0 Or Delay > 320000 Then Exit Sub  ' Más de un minuto no se puede poner
    If DelayIndex < 0 Or DelayIndex > 320000 Then Exit Sub  ' Más de un minuto no se puede poner
      
    ' @ Si no es seteo, comprueba de que sea él, para que otro no le modifique..
    If (Delay > 0 And Mode > 0) Then
    
        Call mStreamer.Streamer_Initial(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Call WriteConsoleMsg(UserIndex, "MODO STREAMER BOT", FontTypeNames.FONTTYPE_INFOGREEN)
       ' Call WriteConsoleMsg(UserIndex, "Tiempo de Warp: " & PonerPuntos(Delay), FontTypeNames.FONTTYPE_INFOGREEN)
        'Call WriteConsoleMsg(UserIndex, "Modo Seleccionado: " & Streamer_Mode_String(Mode), FontTypeNames.FONTTYPE_INFOGREEN)
        Call Streamer_CheckPosition
        'Exit Sub
    ElseIf Delay = 0 And Mode = 0 And DelayIndex = 0 Then
        Call mStreamer.Streamer_Initial(0, 0, 0, 0)
        Call WriteConsoleMsg(UserIndex, "DESACTIVADO MODO STREAMER BOT", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub
    Else
        If StreamerBot.Active <> UserIndex Then Exit Sub
    End If
    
    If Delay > 0 Then
        StreamerBot.Config_TimeWarp = Delay
        Call WriteConsoleMsg(UserIndex, "Tiempo de Warp: " & PonerPuntos(Delay), FontTypeNames.FONTTYPE_INFOGREEN)
    End If
    
    If Mode > 0 Then
        If Mode > eStreamerMode.e_LAST - 1 Then Exit Sub
        
        StreamerBot.Mode = Mode
        Call WriteConsoleMsg(UserIndex, "Modo Seleccionado: " & Streamer_Mode_String(Mode), FontTypeNames.FONTTYPE_INFOGREEN)
    End If

    If DelayIndex > 0 Then
        
        StreamerBot.Config_TimeCanIndex = DelayIndex
        Call WriteConsoleMsg(UserIndex, "Tiempo para buscar a la misma persona: " & PonerPuntos(DelayIndex), FontTypeNames.FONTTYPE_INFOGREEN)
    End If
    
End Sub


Public Function PrepareMessageUpdateBar(ByVal charindex As Integer, _
                                                 ByRef Tipo As eTypeBar, _
                                                 ByVal Min As Long, _
                                                 ByVal max As Long) As String

    Call Writer.WriteInt(ServerPacketID.UpdateBar)
    
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteInt16(charindex)
    
    Call Writer.WriteInt32(Min)
    Call Writer.WriteInt32(max)


End Function

Public Function PrepareMessageUpdateBarTerrain(ByVal X As Integer, _
                                                                        ByVal Y As Integer, _
                                                                        ByRef Tipo As eTypeBar, _
                                                                        ByVal Min As Long, _
                                                                        ByVal max As Long) As String

    Call Writer.WriteInt(ServerPacketID.UpdateBarTerrain)
    
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteInt16(X)
    Call Writer.WriteInt16(Y)
    
    Call Writer.WriteInt32(Min)
    Call Writer.WriteInt32(max)


End Function

Private Sub HandleRequiredLive(ByVal UserIndex As Integer)
    
    If Not Interval_Message(UserIndex) Then Exit Sub

    Call mStreamer.Streamer_RequiredBOT(UserIndex)
    
End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteVelocidadToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.VelocidadToggle)
102     Call Writer.WriteReal32(UserList(UserIndex).Char.speeding)
104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteVelocidadToggle_Err:
        Call Writer.Clear
        '</EhFooter>
End Sub

Public Function PrepareMessageSpeedingACT(ByVal charindex As Integer, _
                                          ByVal speeding As Single)
        '<EhHeader>
        On Error GoTo PrepareMessageSpeedingACT_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SpeedToChar)
102     Call Writer.WriteInt16(charindex)
104     Call Writer.WriteReal32(speeding)
        '<EhFooter>
        Exit Function

PrepareMessageSpeedingACT_Err:
        Call Writer.Clear
       
        '</EhFooter>
End Function


Public Function PrepareMessageMeditateToggle(ByVal charindex As Integer, _
                                             ByVal FX As Integer, _
                                             Optional ByVal X As Integer = 0, _
                                             Optional ByVal Y As Integer = 0, _
                                             Optional ByVal IMeditar As Boolean = True)
        '<EhHeader>
        On Error GoTo PrepareMessageMeditateToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.MeditateToggle)
102     Call Writer.WriteInt16(charindex)
104     Call Writer.WriteInt16(FX)
105     Call Writer.WriteInt16(X)
106     Call Writer.WriteInt16(Y)
          Call Writer.WriteBool(IMeditar)
        '<EhFooter>
        Exit Function

PrepareMessageMeditateToggle_Err:
        Call Writer.Clear
        '</EhFooter>
End Function


Private Sub HandleAcelerationChar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAcelerationChar_Err
        '</EhHeader>
    
        #If Classic = 1 Then
            Exit Sub
        #End If
    
100     With UserList(UserIndex)
        
102         If Not IntervaloPermiteShiftear(UserIndex) Then Exit Sub
        
104         If Not .Stats.MinSta >= (.Stats.MaxSta * 0.3) Then Exit Sub
106         .Counters.BuffoAceleration = 10
108         Call ActualizarVelocidadDeUsuario(UserIndex, True)
        
110         .Stats.MinSta = .Stats.MinSta - (.Stats.MaxSta * 0.3)
112         Call WriteUpdateSta(UserIndex)
        End With
        '<EhFooter>
        Exit Sub

HandleAcelerationChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAcelerationChar " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



Public Sub WriteUpdateUserTrabajo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserIndexInServer_Err
100     Call Writer.WriteInt(ServerPacketID.UpdateUserTrabajo)

104     Call SendData(ToOne, UserIndex, vbNullString)
        '<EhFooter>
        Exit Sub

WriteUserIndexInServer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUserIndexInServer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Private Sub HandleAlquilarComerciante(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo HandleAlquilarComerciante_Err
        '</EhHeader>}
        
        
        
        Dim Tipo As Byte
        
        Tipo = Reader.ReadInt8
        
        
100     With UserList(UserIndex)
        
102         If .flags.TargetNPC = 0 Then
104             Call WriteConsoleMsg(UserIndex, "¡Selecciona la criatura que alquilarás!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
        
106         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.eCommerceChar Then
108             Call WriteChatOverHead(UserIndex, "¡Ey, yo no alquilo mi mercado!", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                Exit Sub
            End If
            
            Exit Sub
            
            If Tipo = 1 Then
110         Call mComerciantes.Commerce_SetNew(.flags.TargetNPC, UserIndex)
            ElseIf Tipo = 2 Then
                Call mComerciantes.Commerce_ViewBalance(.flags.TargetNPC, UserIndex)
            ElseIf Tipo = 3 Then
                Call mComerciantes.Commerce_ReclamarGanancias(.flags.TargetNPC, UserIndex)
            End If
        End With
        '<EhFooter>
        Exit Sub

HandleAlquilarComerciante_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.HandleAlquilarComerciante " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub HandleTirarRuleta(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Dim Mode As Byte
        
        Mode = Reader.ReadInt8
        
        If Mode <> 1 And Mode <> 2 Then Exit Sub
        
       ' Call mRuleta.Ruleta_Tirada(UserIndex, Mode)
    
        
    
    End With
End Sub

Private Sub HandleLotteryNew(ByVal UserIndex As Integer)
    
    Dim TempLottery As tLottery
    
    TempLottery.Name = Reader.ReadString8
    TempLottery.Desc = Reader.ReadString8
    TempLottery.DateFinish = Reader.ReadString8
    TempLottery.PrizeChar = Reader.ReadString8
    TempLottery.PrizeObj = Reader.ReadInt16
    TempLottery.PrizeObjAmount = Reader.ReadInt16
    
    If Len(TempLottery.Name) <= 0 Then Exit Sub
    If Len(TempLottery.Desc) <= 0 Then Exit Sub
    If Len(TempLottery.PrizeChar) <= 0 And TempLottery.PrizeObj <= 0 Then Exit Sub
    If TempLottery.PrizeObj > 0 And TempLottery.PrizeObjAmount <= 0 Then Exit Sub

    If Not EsGmPriv(UserIndex) Then Exit Sub

    Call mLottery.Lottery_New(TempLottery)
End Sub

Public Sub WriteTournamentList(ByVal UserIndex As Integer)
        On Error GoTo WriteTournamentList_Err
100
        Call Writer.WriteInt(ServerPacketID.TournamentList)


        Dim A As Long, B As Long
        
        For A = 1 To MAX_EVENT_SIMULTANEO
            With Events(A)
                If .Name <> vbNullString Then
                    Call Writer.WriteBool(True)
                    
                    Call Writer.WriteString8(.Name)
                    Call Writer.WriteInt8(.config(eConfigEvent.eFuegoAmigo))
                    Call Writer.WriteInt8(.LimitRound)
                    Call Writer.WriteInt8(.LimitRoundFinal)
                    Call Writer.WriteInt16(.PrizePoints)
                    Call Writer.WriteInt8(.LvlMin)
                    Call Writer.WriteInt8(.LvlMax)
                    
                    For B = 1 To NUMCLASES
                        Call Writer.WriteInt8(.AllowedClasses(B))
                    Next B
                    
                    Call Writer.WriteInt16(.InscriptionGld)
                    Call Writer.WriteInt16(.InscriptionEldhir)
                    
                    Call Writer.WriteInt16(.PrizeGld)
                    Call Writer.WriteInt16(.PrizeEldhir)
                    Call Writer.WriteInt16(.PrizeObj.ObjIndex)
                    Call Writer.WriteInt16(.PrizeObj.Amount)
                    
                    Call Writer.WriteInt8(.config(eConfigEvent.eCascoEscudo))
                    
                    Call Writer.WriteInt8(.config(eConfigEvent.eResu))
                    Call Writer.WriteInt8(.config(eConfigEvent.eInvisibilidad))
                    Call Writer.WriteInt8(.config(eConfigEvent.eOcultar))
                    Call Writer.WriteInt8(.config(eConfigEvent.eInvocar))
                Else
                    Call Writer.WriteBool(False)
                End If
            End With
        Next A



104     Call SendData(ToOne, UserIndex, vbNullString)

        Exit Sub

WriteTournamentList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteTournamentList " & _
               "at line " & Erl

End Sub

Public Sub WriteStatsUser(ByVal UserIndex As Integer, ByRef IUser As User)
     On Error GoTo WriteStatsUser_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser)

    Dim A As Long
    
    With IUser
        Call Writer.WriteString8(.Name)
        Call Writer.WriteInt8(.Clase)
        Call Writer.WriteInt8(.Raza)
        Call Writer.WriteInt8(.Genero)
        Call Writer.WriteInt8(.Stats.Elv)
        Call Writer.WriteInt32(.Stats.Exp)
        Call Writer.WriteInt32(.Stats.Elu)
        
        Call Writer.WriteInt8(.Blocked)
        Call Writer.WriteInt32(.BlockedHasta)
        
        Call Writer.WriteInt32(.Stats.Gld)
        Call Writer.WriteInt32(.Stats.Eldhir)
        Call Writer.WriteInt32(.Stats.Points)
        
        Call Writer.WriteInt16(.Faction.FragsOther)
        Call Writer.WriteInt16(.Faction.FragsCiu)
        Call Writer.WriteInt16(.Faction.FragsCri)
        
        Call Writer.WriteInt16(.Pos.Map)
        Call Writer.WriteInt8(.Pos.X)
        Call Writer.WriteInt8(.Pos.Y)
        
        Call Writer.WriteInt16(.Stats.MaxHp)
    End With
    

    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser " & _
               "at line " & Erl
End Sub

' # Inventario de un personaje
Public Sub WriteStatsUser_Inventory(ByVal UserIndex As Integer, ByRef IUser As Inventario)
     On Error GoTo WriteStatsUser_Inventory_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Inventory)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.NroItems)
        
        If .NroItems > 0 Then
            For A = 1 To .NroItems
                Call Writer.WriteInt16(.Object(A).ObjIndex)
                Call Writer.WriteInt16(.Object(A).Amount)
                Call Writer.WriteInt8(.Object(A).Equipped)
            Next A
        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Inventory_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Inventory " & _
               "at line " & Erl
End Sub
' # Banco de un personaje
Public Sub WriteStatsUser_Bank(ByVal UserIndex As Integer, ByRef IUser As BancoInventario)
     On Error GoTo WriteStatsUser_Inventory_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Bank)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.NroItems)
        
        If .NroItems > 0 Then
            For A = 1 To .NroItems
                Call Writer.WriteInt16(.Object(A).ObjIndex)
                Call Writer.WriteInt16(.Object(A).Amount)
                Call Writer.WriteInt8(.Object(A).Equipped)
            Next A
        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Inventory_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Inventory " & _
               "at line " & Erl
End Sub
' # Hechizos de un personaje
Public Sub WriteStatsUser_Spells(ByVal UserIndex As Integer, ByRef IUser() As Integer)
     On Error GoTo WriteStatsUser_Spells_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Spells)

    Dim A As Long

    For A = LBound(IUser) To UBound(IUser)
        Call Writer.WriteInt16(IUser(A))
    Next A

    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Spells_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Spells " & _
               "at line " & Erl
End Sub

' # Habilidades de un personaje
Public Sub WriteStatsUser_Skills(ByVal UserIndex As Integer, ByRef IUser() As Integer)
    On Error GoTo WriteStatsUser_Skills_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Skills)
    
    Dim A As Long

    For A = LBound(IUser) To UBound(IUser)
        Call Writer.WriteInt16(IUser(A))
    Next A
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Skills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Skills " & _
               "at line " & Erl
End Sub
' # Bonificaciones de un personaje por tiempo de duración
Public Sub WriteStatsUser_Bonus(ByVal UserIndex As Integer, ByRef IUser As UserStats)
     On Error GoTo WriteStatsUser_Bonos_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Bonos)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.BonusLast)
        
        If .BonusLast > 0 Then
            For A = 1 To .BonusLast
                With .Bonus(A)
                    Call Writer.WriteInt8(.Tipo)
                    Call Writer.WriteInt(.Value)
                    Call Writer.WriteInt(.Amount)
                    Call Writer.WriteInt(.DurationSeconds)
                    Call Writer.WriteString8(.DurationDate)
                End With
            Next A
        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Bonos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Bonos_Err " & _
               "at line " & Erl
End Sub

' # Penas de un personaje por tiempo de duración
Public Sub WriteStatsUser_Penas(ByVal UserIndex As Integer, ByRef IUser As User)
     On Error GoTo WriteStatsUser_Penas_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Penas)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt16(.Counters.Pena)
        
        
        Call Writer.WriteInt8(.PenasLast)
        
        ' # Cargar Penas
        If .PenasLast > 0 Then
            For A = 1 To .PenasLast
                Call Writer.WriteString8(.Penas(A))
            Next A
        End If
    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Penas_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Penas " & _
               "at line " & Erl
End Sub
' # Skins de un personaje
Public Sub WriteStatsUser_Skins(ByVal UserIndex As Integer, ByRef IUser As tSkins)
     On Error GoTo WriteStatsUser_Skins_Err
100
    Call Writer.WriteInt(ServerPacketID.StatsUser_Skins)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.Last)
        
        If .Last > 0 Then
            For A = 1 To .Last
                Call Writer.WriteInt16(.ObjIndex(A))
            Next A
        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Skins_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteStatsUser_Skins " & _
               "at line " & Erl
End Sub

Public Sub WriteUpdateClient(ByVal UserIndex As Integer)
     On Error GoTo WriteStatsUser_Err
100
    Call Writer.WriteInt(ServerPacketID.UpdateClient)
    
    Call SendData(ToOne, UserIndex, vbNullString, , True)

    Exit Sub

WriteStatsUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Protocol.WriteUpdateClient " & _
               "at line " & Erl
End Sub

Private Sub HandleCastleInfo(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    
    Dim A As Long
    Dim Text As String
    
    Dim CastleIndex As Byte
    
    CastleIndex = Reader.ReadInt8
    
    If CastleIndex > 4 Then Exit Sub
    
    If CastleIndex = 0 Then
        If Not Interval_Message(UserIndex) Then Exit Sub
        
        For A = 1 To CastleLast
            With Castle(A)
                Text = .Name & "» " & .Desc & IIf(.GuildIndex > 0, " (Conquistado por: " & .GuildName & ")", " NO está conquistado.")
                
                Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_USERBRONCE)
                
            End With
        Next A
        
        Call WriteConsoleMsg(UserIndex, "BONUS 10% EXP+ORO» " & IIf(CastleBonus > 0, "(Obtenido por: " & GuildsInfo(CastleBonus).Name & ")", "Ningún clan es poseedor."), FontTypeNames.FONTTYPE_USERPLATA)

        Call WriteConsoleMsg(UserIndex, "Utiliza los comandos /NORTE /SUR /ESTE /OESTE una vez que seas poseedor del Castillo.", FontTypeNames.FONTTYPE_USERGOLD)
    Else
        Call mCastle.Castle_Travel(UserIndex, CastleIndex)
    
    End If
    
    Exit Sub
ErrHandler:
    
End Sub

Private Sub HandleRequiredStatsUser(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim Tipo As Byte
    Dim Name As String
    Dim IUser As User
    Dim tUser As Integer
    
    Tipo = Reader.ReadInt8
    Name = Reader.ReadString8
    
    If Tipo < 0 Then Exit Sub
    
    ' # Chequea el intervalo con el que lo hace
    If Not Interval_Packet500(UserIndex) Then
        Call Logs_Security(eSecurity, eAntiHack, "El requerimiento de Stats está siendo alto por parte de " & UserList(UserIndex).Account.Email)
        Exit Sub
    End If
    
    ' # No existe el personaje
    If Not PersonajeExiste(Name) Then
        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub
    End If
    
    ' # Está online
    tUser = NameIndex(Name)
    
    If tUser > 0 Then
        IUser = UserList(tUser)
        
        ' Información confidencial
        If MapInfo(IUser.Pos.Map).Pk Then
            IUser.Pos.Map = 0
            IUser.Invent.NroItems = 0
        End If
        
        If EsGm(tUser) Then Exit Sub
    Else
        If EsDios(Name) Or EsSemiDios(Name) Or EsAdmin(Name) Then Exit Sub
        IUser = Load_UserList_Offline(Name)         ' # Cargamos el personaje offline
        
    End If
    
    Select Case Tipo
        Case 0 ' Inventario
            Call WriteStatsUser_Inventory(UserIndex, IUser.Invent)
            
        Case 1 ' Spells
            Call WriteStatsUser_Spells(UserIndex, IUser.Stats.UserHechizos)
            
        Case 2 ' Boveda
            Call WriteStatsUser_Bank(UserIndex, IUser.BancoInvent)
            
        Case 3 ' Skills
            Call WriteStatsUser_Skills(UserIndex, IUser.Stats.UserSkills)
            
        Case 4 ' Bonus
            Call WriteStatsUser_Bonus(UserIndex, IUser.Stats)
            
        Case 5 ' Penas
            Call WriteStatsUser_Penas(UserIndex, IUser)
            
        Case 6 ' Skins
            Call WriteStatsUser_Skins(UserIndex, IUser.Skins)
            
        Case 7 ' Logros
            ' # Proximamente
            
        Case 197 ' Formulario principal
            Call WriteStatsUser(UserIndex, IUser)
            
            
    End Select
    
    ' # Numeros que puede solicitar
    'eInventory = 0
    'eSpells = 1
    'eBank = 2
    'eAbilities = 3
    'eBonus = 4
    'ePenas = 5
    'eSkins = 6
    'eLogros = 7
    
    
ErrHandler:
    Exit Sub
    
End Sub

