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
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Public Const CustomPath As String = "custom.dat"

Public AdminMsg         As Byte

Public InfoMsg          As Byte

Public GuildMsg         As Byte

Public PartyMsg         As Byte

Public CombateMsg       As Byte

Public TrabajoMsg       As Byte

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

Public Enum eConsoleType

    General = 0
    Acciones = 1
    Agrupaciones = 2
    Custom = 3

End Enum

Private Type tFont

    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean

End Type

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
    StopWave
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
    Castle
    RequiredStatsUser
    
End Enum

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
    FONTTYPE_STREAMER
    FONTTYPE_RMSG
    FONTTYPE_RACHAS
End Enum

Public FontTypes(39) As tFont

' Initializes the fonts array

Public Sub InitFonts()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With FontTypes(FontTypeNames.FONTTYPE_RACHAS)
        .red = 47
        .green = 234
        .blue = 170
        .bold = 222
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_RMSG)
        .red = 180
        .green = 255
        .blue = 170
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_STREAMER)
        .red = 200
        .green = 150
        .blue = 245
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ADMIN)
        .red = 250
        .green = 180
        .blue = 50
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 32
        .green = 51
        .blue = 223
        .bold = 1
        .italic = 1
        ''
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 249
        .green = 244
        .blue = 204
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFORED)
        .red = 249
        .green = 202
        .blue = 202
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOGREEN)
        .red = 208
        .green = 249
        .blue = 202
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 130
        .green = 130
        .blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .red = 255
        .green = 180
        .blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 228
        .green = 199
        .blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .red = 130
        .green = 130
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .green = 200
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 255
        .green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .red = 255
        .green = 255
        .blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .red = 30
        .green = 255
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .red = 30
        .green = 150
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .red = 250
        .green = 250
        .blue = 150
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EVENT)
        .red = 9
        .green = 202
        .blue = 80
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_USERGOLD)
        .red = 250
        .green = 200
        .blue = 10
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_USERPLATA)
        .red = 130
        .green = 130
        .blue = 130
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_USERBRONCE)
        .red = 255
        .green = 166
        .blue = 0
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM)
        .red = 247
        .green = 170
        .blue = 15
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ANGEL)
        .red = 130
        .green = 247
        .blue = 247
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DEMONIO)
        .red = 160
        .green = 17
        .blue = 17
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
        .red = 175
        .green = 250
        .blue = 250
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CRITICO)
        .red = 115
        .green = 227
        .blue = 215
        .bold = True
    End With

    With FontTypes(FontTypeNames.FONTTYPE_INFORETOS)
        .red = 245
        .green = 158
        .blue = 75
        .bold = False
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INVASION)
        .red = 243
        .green = 180
        .blue = 243
        .bold = True
        .italic = True
    End With
    
        
    With FontTypes(FontTypeNames.FONTTYPE_PODER)
        .red = 23
        .green = 248
        .blue = 163
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DESAFIOS)
        .red = 7
        .green = 250
        .blue = 250
        .bold = False
    End With
    
    
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()

        '<EhHeader>
        On Error GoTo HandleIncomingData_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        Dim Packet As Long

100     Packet = Reader.ReadInt
    
102     Select Case Packet
            
            Case ServerPacketID.UpdateClient
                Call HandleUpdateClient
                
            Case ServerPacketID.StatsUser
                Call HandleStatsUser
                
            Case ServerPacketID.StatsUser_Inventory
                Call HandleStatsUser_Inventory
                
            Case ServerPacketID.StatsUser_Bank
                Call HandleStatsUser_Bank
                
            Case ServerPacketID.StatsUser_Spells
                Call HandleStatsUser_Spells
                
            Case ServerPacketID.StatsUser_Skills
                Call HandleStatsUser_Skills
                
            Case ServerPacketID.StatsUser_Bonos
                Call HandleStatsUser_Bonus
                
            Case ServerPacketID.StatsUser_Penas
                Call HandleStatsUser_Penas
                
            Case ServerPacketID.StatsUser_Skins
                Call HandleStatsUser_Skins
                
            Case ServerPacketID.StatsUser_Logros
                'Call HandleStatsUser_Logros
                
            Case ServerPacketID.TournamentList
                Call HandleTournamentList
                
            Case ServerPacketID.UpdateUserTrabajo
                Call HandleUpdateUserTrabajo
                
            Case ServerPacketID.VelocidadToggle
                Call HandleVelocidadToggle
            
            Case ServerPacketID.SpeedToChar
                Call HandleSpeedToChar
            
            Case ServerPacketID.RequiredMoveChar
                Call HandleRequiredMoveChar
                
            Case ServerPacketID.UpdateDataSkin
104             Call HandleUpdateDataSkin
            
106         Case ServerPacketID.UpdateFinishQuest
108             Call HandleUpdateFinishQuest
            
110         Case ServerPacketID.SendShopChars
112             Call HandleSendShopChars
            
118         Case ServerPacketID.UpdateMeditation
120             Call HandleUpdateMeditation
            
122         Case ServerPacketID.UpdateEvento
124             Call HandleModoEvento
            
126         Case ServerPacketID.UpdateOnline
128             Call HandleUpdateOnline
            
130         Case ServerPacketID.SendIntervals
132             Call HandleReceiveIntervals
            
134         Case ServerPacketID.UpdateEffectPoison
136             Call HandleUpdateEffectPoison
            
138         Case ServerPacketID.Render_CountDown
140             Call HandleRender_CountDown
            
142         Case ServerPacketID.Connected
144             Call HandleConnectedMessage
            
146         Case ServerPacketID.LoggedRemoveChar
148             Call HandleLoggedRemoveChar
            
150         Case ServerPacketID.loggedaccount
152             Call HandleLoggedAccount

        
154         Case ServerPacketID.AccountInfo
156             Call HandleAccountInfo
            
158         Case ServerPacketID.LoggedAccount_DataChar
160             Call HandleLoggedAccount_DataChar
            
162         Case ServerPacketID.Mercader_ListOffer
164             Call HandleMercader_ListOffer
            
166         Case ServerPacketID.Mercader_ListInfo
168             Call HandleMercader_ListChar
            
170         Case ServerPacketID.Mercader_List
172             Call HandleMercader_List
        
174         Case ServerPacketID.MiniMap_InfoCriature
176             Call HandleMiniMap_InfoCriature
            
178         Case ServerPacketID.Fight_PanelAccept
180             Call HandleFight_PanelInvitation
            
182         Case ServerPacketID.Guild_InfoUsers
184             Call HandleGuild_InfoUsers
            
186         Case ServerPacketID.Guild_Info
188             Call HandleGuild_Info
            
190         Case ServerPacketID.Guild_List
192             Call HandleGuild_List
            
194         Case ServerPacketID.UpdateGroupIndex
196             Call HandleUpdateGroupIndex
            
198         Case ServerPacketID.UpdateInfoIntervals
200             Call HandleUpdateInfoIntervals
            
202         Case ServerPacketID.UpdateListSecurity
204             Call HandleUpdateListSecurity
    
206         Case ServerPacketID.UpdateControlPotas
208             Call HandleUpdateControlPotas
                
            Case ServerPacketID.UpdateBar
                Call HandleUpdateBar
                
            Case ServerPacketID.UpdateBarTerrain
                Call HandleUpdateBarTerrain
            
210         Case ServerPacketID.ClickVesA
212             Call HandleClickVesA
          
214         Case ServerPacketID.CreateDamage
216             Call HandleCreateDamage
            
218         Case ServerPacketID.SolicitaCapProc
220             Call HandleSolicitaCapProc
            
222         Case ServerPacketID.UserInEvent
224             Call HandleUserInEvent
            
226         Case ServerPacketID.SendInfoRetos
228             Call HandleSendRetos
            
238         Case ServerPacketID.MontateToggle
240             Call HandleMontateToggle
            
242         Case ServerPacketID.GroupPrincipal
244             Call HandleGroupPrincipal
            
246         Case ServerPacketID.GroupUpdateExp
248             Call HandleGroupUpdateExp
                
250         Case ServerPacketID.logged                  ' LOGGED
252             Call HandleLogged
        
254         Case ServerPacketID.RemoveDialogs           ' QTDL
256             Call HandleRemoveDialogs
        
258         Case ServerPacketID.RemoveCharDialog        ' QDL
260             Call HandleRemoveCharDialog
        
262         Case ServerPacketID.NavigateToggle          ' NAVEG
264             Call HandleNavigateToggle
        
266         Case ServerPacketID.Disconnect              ' FINOK
268             Call HandleDisconnect
        
270         Case ServerPacketID.CommerceEnd             ' FINCOMOK
272             Call HandleCommerceEnd
            
274         Case ServerPacketID.CommerceChat
276             Call HandleCommerceChat
        
278         Case ServerPacketID.BankEnd                 ' FINBANOK
280             Call HandleBankEnd
        
282         Case ServerPacketID.CommerceInit            ' INITCOM
284             Call HandleCommerceInit
        
286         Case ServerPacketID.BankInit                ' INITBANCO
288             Call HandleBankInit
        
294         Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
296             Call HandleUserCommerceInit
        
298         Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
300             Call HandleUserCommerceEnd
            
302         Case ServerPacketID.UserOfferConfirm
304             Call HandleUserOfferConfirm
        
306         Case ServerPacketID.UpdateSta               ' ASS
308             Call HandleUpdateSta
        
310         Case ServerPacketID.UpdateMana              ' ASM
312             Call HandleUpdateMana
        
314         Case ServerPacketID.UpdateHP                ' ASH
316             Call HandleUpdateHP
        
318         Case ServerPacketID.UpdateGold              ' ASG
320             Call HandleUpdateGold
            
322         Case ServerPacketID.UpdateDsp
324             Call HandleUpdateDsp
            
326         Case ServerPacketID.UpdateBankGold
328             Call HandleUpdateBankGold

330         Case ServerPacketID.UpdateExp               ' ASE
332             Call HandleUpdateExp
            
334         Case ServerPacketID.ChangeMap               ' CM
336             Call HandleChangeMap
        
338         Case ServerPacketID.PosUpdate               ' PU
340             Call HandlePosUpdate
        
342         Case ServerPacketID.ChatOverHead            ' ||
344             Call HandleChatOverHead
        
346         Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
348             Call HandleConsoleMessage
        
350         Case ServerPacketID.ShowMessageBox          ' !!
352             Call HandleShowMessageBox
        
354         Case ServerPacketID.UserIndexInServer       ' IU
356             Call HandleUserIndexInServer
        
358         Case ServerPacketID.UserCharIndexInServer   ' IP
360             Call HandleUserCharIndexInServer
        
362         Case ServerPacketID.CharacterCreate         ' CC
364             Call HandleCharacterCreate
        
366         Case ServerPacketID.CharacterChangeHeading
368             Call HandleCharacterChangeHeading
        
370         Case ServerPacketID.CharacterRemove         ' BP
372             Call HandleCharacterRemove
        
374         Case ServerPacketID.CharacterChangeNick
376             Call HandleCharacterChangeNick
            
378         Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
380             Call HandleCharacterMove
        
382         Case ServerPacketID.CharacterAttackMovement
384             Call HandleCharacterAttackMovement
            
386         Case ServerPacketID.CharacterAttackNpc
388             Call HandleCharacterAttackNpc
            
390         Case ServerPacketID.ForceCharMove
392             Call HandleForceCharMove
        
394         Case ServerPacketID.CharacterChange         ' CP
396             Call HandleCharacterChange
        
398         Case ServerPacketID.ObjectCreate            ' HO
400             Call HandleObjectCreate
        
402         Case ServerPacketID.ObjectDelete            ' BO
404             Call HandleObjectDelete
        
406         Case ServerPacketID.BlockPosition           ' BQ
408             Call HandleBlockPosition
        
410         Case ServerPacketID.PlayMusic               ' TM
412             Call HandlePlayMusic
        
414         Case ServerPacketID.PlayWave              ' TW
416             Call HandlePlayWave
              
            Case ServerPacketID.StopWave
                Call HandleStopWave
                
418         Case ServerPacketID.PauseToggle             ' BKW
420             Call HandlePauseToggle
        
422         Case ServerPacketID.CreateFX                ' CFX
424             Call HandleCreateFX
            
426         Case ServerPacketID.CreateFXMap
428             Call HandleCreateFXMap
        
430         Case ServerPacketID.UpdateUserStats         ' EST
432             Call HandleUpdateUserStats
        
434         Case ServerPacketID.ChangeInventorySlot     ' CSI
436             Call HandleChangeInventorySlot
        
438         Case ServerPacketID.ChangeBankSlot          ' SBO
440             Call HandleChangeBankSlot
            
442         Case ServerPacketID.ChangeBankSlot_Account
444             Call HandleChangeBankSlot_Account
        
446         Case ServerPacketID.ChangeSpellSlot         ' SHS
448             Call HandleChangeSpellSlot
        
450         Case ServerPacketID.Atributes               ' ATR
452             Call HandleAtributes
        
454         Case ServerPacketID.RestOK                  ' DOK
456             Call HandleRestOK
        
458         Case ServerPacketID.ErrorMsg                ' ERR
460             Call HandleErrorMessage
        
462         Case ServerPacketID.Blind                   ' CEGU
464             Call HandleBlind
        
466         Case ServerPacketID.Dumb                    ' DUMB
468             Call HandleDumb

470         Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
472             Call HandleChangeNPCInventorySlot
        
474         Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
476             Call HandleUpdateHungerAndThirst
        
478         Case ServerPacketID.MiniStats               ' MEST
480             Call HandleMiniStats
        
482         Case ServerPacketID.LevelUp                 ' SUNI
484             Call HandleLevelUp

486         Case ServerPacketID.SetInvisible            ' NOVER
488             Call HandleSetInvisible

490         Case ServerPacketID.MeditateToggle          ' MEDOK
492             Call HandleMeditateToggle
        
494         Case ServerPacketID.BlindNoMore             ' NSEGUE
496             Call HandleBlindNoMore
        
498         Case ServerPacketID.DumbNoMore              ' NESTUP
500             Call HandleDumbNoMore
        
502         Case ServerPacketID.SendSkills              ' SKILLS
504             Call HandleSendSkills
        
506         Case ServerPacketID.ParalizeOK              ' PARADOK
508             Call HandleParalizeOK
        
510         Case ServerPacketID.ShowUserRequest         ' PETICIO
512             Call HandleShowUserRequest
        
514         Case ServerPacketID.TradeOK                 ' TRANSOK
516             Call HandleTradeOK
        
518         Case ServerPacketID.BankOK                  ' BANCOOK
520             Call HandleBankOK
        
522         Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
524             Call HandleChangeUserTradeSlot
        
526         Case ServerPacketID.Pong
528             Call HandlePong
        
530         Case ServerPacketID.UpdateTagAndStatus
532             Call HandleUpdateTagAndStatus
        
                '*******************
                'GM messages
                '*******************
534         Case ServerPacketID.SpawnList               ' SPL
536             Call HandleSpawnList
            
538         Case ServerPacketID.ShowDenounces
540             Call HandleShowDenounces
            
542         Case ServerPacketID.RecordDetails
544             Call HandleRecordDetails
            
546         Case ServerPacketID.RecordList
548             Call HandleRecordList
        
550         Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
552             Call HandleShowGMPanelForm
        
554         Case ServerPacketID.UserNameList            ' LISTUSU
556             Call HandleUserNameList
        
558         Case ServerPacketID.UpdateStrenghtAndDexterity
560             Call HandleUpdateStrenghtAndDexterity
            
562         Case ServerPacketID.UpdateStrenght
564             Call HandleUpdateStrenght
            
566         Case ServerPacketID.UpdateDexterity
568             Call HandleUpdateDexterity
            
570         Case ServerPacketID.AddSlots
572             Call HandleAddSlots

574         Case ServerPacketID.MultiMessage
576             Call HandleMultiMessage
            
578         Case ServerPacketID.CancelOfferItem
580             Call HandleCancelOfferItem
            
582         Case ServerPacketID.ShowMenu
584             Call HandleShowMenu
            
586         Case ServerPacketID.StrDextRunningOut
588             Call HandleStrDextRunningOut
            
590         Case ServerPacketID.ChatPersonalizado
592             Call HandleChatPersonalizado
            
598         Case ServerPacketID.RenderConsole
600             Call HandleRenderConsole
            
602         Case ServerPacketID.ViewListQuest
604             Call HandleViewListQuest
        
606         Case ServerPacketID.UpdateUserDead
608             Call HandleUpdateUserDead
        
610         Case ServerPacketID.QuestInfo
612             Call HandleQuestData
        
614         Case ServerPacketID.UpdateGlobalCounter
616             Call HandleUpdateGlobalCounter
            
618         Case ServerPacketID.SendInfoNpc
620             Call HandleSendInfoNpc
        
622         Case ServerPacketID.UpdatePosGuild
624             Call HandleUpdatePosGuild
            
626         Case ServerPacketID.UpdateLevelGuild
628             Call HandleUpdateLevelGuild
            
630         Case ServerPacketID.UpdateStatusMAO
632             Call HandleUpdateStatusMAO

634         Case Else

                'ERROR : Abort!
                Exit Sub
       
        End Select

        '<EhFooter>
        Exit Sub

HandleIncomingData_Err:
        LogError err.Description & vbCrLf & "in ARGENTUM.Protocol.HandleIncomingData " & "at line " & Erl
               
        Resume Next

        '</EhFooter>
End Sub

Public Sub HandleMultiMessage()

    '***************************************************
    'Author: Unknown
    'Last Modification: 11/16/2010
    ' 09/28/2010: C4b3z0n - Ahora se le saco la "," a los minutos de distancia del /hogar, ya que a veces quedaba "12,5 minutos y 30segundos"
    ' 09/21/2010: C4b3z0n - Now the fragshooter operates taking the screen after the change of killed charindex to ghost only if target charindex is visible to the client, else it will take screenshot like before.
    ' 11/16/2010: Amraphen - Recoded how the FragShooter works.
    '***************************************************
    Dim BodyPart As Byte

    Dim Daño As Integer
    
    Select Case Reader.ReadInt
        
        Case eMessages.DontSeeAnything
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_NO_VES_NADA_INTERESANTE, 65, 190, 156, False, False, True)
        
        Case eMessages.NPCSwing
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.NPCKillUser
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldUser
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldOther
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserSwing
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.SafeModeOn
                #If ModoBig = 0 Then
                    FrmMain.imgButton(6).Picture = LoadPicture(DirInterface & "main\imgSeg.jpg")
                #Else
                    FrmMain.imgButton(6).Picture = LoadPicture(DirInterface & "main\imgSegBig.jpg")
                #End If
                
                Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
            
            IsSeguro = True

        Case eMessages.SafeModeOff
            FrmMain.imgSeg.Picture = Nothing
            
            FrmMain.imgButton(6).Picture = Nothing
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
             
            IsSeguro = False
            
        Case eMessages.ResuscitationSafeOff
            FrmMain.imgResu.Picture = Nothing
            FrmMain.imgButton(8).Picture = Nothing
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, True)


        Case eMessages.ResuscitationSafeOn
                Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True)
                
                #If ModoBig = 0 Then
                    FrmMain.imgButton(8).Picture = LoadPicture(DirInterface & "main\imgResu.jpg")
                #Else
                    FrmMain.imgButton(8).Picture = LoadPicture(DirInterface & "main\imgResuBig.jpg")
                #End If
                
        Case eMessages.DragSafeOn

                Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_DRAG_ACTIVADO, 0, 255, 0, True, False, True)
                
                #If ModoBig = 0 Then
                    FrmMain.imgButton(7).Picture = LoadPicture(DirInterface & "main\ImgDrag.jpg")
                #Else
                    FrmMain.imgButton(7).Picture = LoadPicture(DirInterface & "main\ImgDragBig.jpg")
                #End If
            
        Case eMessages.DragSafeOff
            FrmMain.imgDrag.Picture = Nothing
            FrmMain.imgButton(7).Picture = Nothing
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_DRAG_DESACTIVADO, 255, 0, 0, True, False, True)

        Case eMessages.NobilityLost
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
        
        Case eMessages.CantUseWhileMeditating
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
        
        Case eMessages.NPCHitUser

            Select Case Reader.ReadInt()

                Case bCabeza
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(Reader.ReadInt() & "!!"), 255, 0, 0, True, False, True)

            End Select
        
        Case eMessages.UserHitNPC
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(Reader.ReadInt()) & MENSAJE_2, 255, 0, 0, True, False, True)
        
        Case eMessages.UserAttackedSwing
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & CharList(Reader.ReadInt()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserHittedByUser

            Dim AttackerName As String
            
            AttackerName = GetRawName(CharList(Reader.ReadInt()).Nombre)
            BodyPart = Reader.ReadInt()
            Daño = Reader.ReadInt()
            
            Select Case BodyPart

                Case bCabeza
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)

            End Select
        
        Case eMessages.UserHittedUser

            Dim VictimName As String
            
            VictimName = GetRawName(CharList(Reader.ReadInt()).Nombre)
            BodyPart = Reader.ReadInt()
            Daño = Reader.ReadInt()
            
            Select Case BodyPart

                Case bCabeza
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)

            End Select
        
        Case eMessages.WorkRequestTarget
            UsingSkill = Reader.ReadInt()
            
            FrmMain.MousePointer = 2
            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorSpell, IDC_CROSS)

            Select Case UsingSkill

                Case Magia
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                
                Case Robar
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
  
                Case Proyectiles
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
                
                Case TeleportInvoker
                    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_TELEPORT_INVOKER, 100, 100, 120, 0, 0)
            End Select

        Case eMessages.HaveKilledUser

            Dim KilledUser As Integer

            Dim Exp        As Long
            
            KilledUser = Reader.ReadInt
            Exp = Reader.ReadInt
            
            Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & CharList(KilledUser).Nombre & MENSAJE_22, 255, 0, 0, True, False)
            Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & Exp & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
            
            'Sacamos un screenshot si está activado el FragShooter:
            'If ClientSetup.bKill And ClientSetup.bActive Then
            'If Exp \ 2 > ClientSetup.byMurderedLevel Then
            'FragShooterNickname = CharList(KilledUser).Nombre
            'FragShooterKilledSomeone = True
                    
            'FragShooterCapturePending = True
            'End If
            'End If
            
        Case eMessages.UserKill

            Dim KillerUser As Integer
            
            KillerUser = Reader.ReadInt
            
            Call ShowConsoleMsg(CharList(KillerUser).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)
                
        Case eMessages.EarnExp
            'Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & Reader.Readint & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        
        Case eMessages.GoHome

            Dim Distance As Byte

            Dim Hogar    As String

            Dim tiempo   As Integer

            Dim msg      As String
            
            Distance = Reader.ReadInt
            tiempo = Reader.ReadInt
            Hogar = Reader.ReadString8
            
            If tiempo >= 60 Then
                If tiempo Mod 60 = 0 Then
                    msg = tiempo / 60 & " minutos."
                Else
                    msg = CInt(tiempo \ 60) & " minutos y " & tiempo Mod 60 & " segundos."  'Agregado el CInt() asi el número no es con , [C4b3z0n - 09/28/2010]

                End If

            Else
                msg = tiempo & " segundos."

            End If
            
            Call ShowConsoleMsg("Viajarás a la Ciudad Principal y el viaje durará " & msg, 255, 0, 0, True)
            Traveling = True
        
        Case eMessages.FinishHome
            Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
            Traveling = False
        
        Case eMessages.CancelGoHome
            Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
            Traveling = False
        
        
        Case eMessages.ModoStreamOn
            Call ShowConsoleMsg(MENSAJE_MODOSTREAM_ON1, 200, 150, 245, False)
            Call ShowConsoleMsg(MENSAJE_MODOSTREAM_ON2, 200, 150, 245, True)
            Call ShowConsoleMsg(MENSAJE_MODOSTREAM_ON3, 180, 150, 150, True)
            
        Case eMessages.ModoStreamOff
            Call ShowConsoleMsg(MENSAJE_MODOSTREAM_OFF, 200, 250, 200, True)
    End Select

End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    ' Variable initialization
    UserClase = Reader.ReadInt8
    UserRaza = Reader.ReadInt8
    UserSexo = Reader.ReadInt8()
    Account.CharsAmount = Reader.ReadInt8
    Account.Gld = Reader.ReadInt32
    
    Nombres = True

    'Set connected state
    Call SetConnected

    ' Call DibujarMiniMapa
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check if the packet is complete
    
    Call Dialogos.RemoveDialog(Reader.ReadInt())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserNavegando = Not UserNavegando
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    ResetAllInfo (Reader.ReadBool)
    
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'Hide form
    Unload frmComerciar
    
    Set InvComUser = Nothing
    Set InvComNpc = Nothing
    
    'Reset vars
    Comerciando = False
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Unload frmBancoObj
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Comerciando = False
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I As Long
    
    NpcName = Reader.ReadString8
    QuestLast = Reader.ReadInt8
    
    If QuestLast > 0 Then
        Call Reader.ReadSafeArrayInt8(QuestNpc)
    End If
    
    Set InvComNpc = New clsGrapchicalInventory
    Set InvComUser = New clsGrapchicalInventory

    ' Initialize commerce inventories
    Call InvComUser.Initialize(frmComerciar.picInvUser, Inventario.MaxObjs, Inventario.MaxObjs, eCaption.Comercio_User, , , , , , , , , , True)
    Call InvComNpc.Initialize(frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS, MAX_NPC_INVENTORY_SLOTS, eCaption.Comercio_Npc, , , , , , False, , , , True, , True)

    'Fill user inventory
    For I = 1 To MAX_INVENTORY_SLOTS

        'If Inventario.OBJIndex(i) <> 0 Then
        With Inventario
            Call InvComUser.SetItem(I, .ObjIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .ValorAzul(I), .CanUse(I), .MinHitMag(I), .MaxHitMag(I), .MinDefMag(I), .MaxDefMag(I))
        End With

        'End If
    Next I
    
    ' Fill Npc inventory
    For I = 1 To MAX_NPC_INVENTORY_SLOTS

        'If NPCInventory(i).OBJIndex <> 0 Then
        With NPCInventory(I)
            Call InvComNpc.SetItem(I, .ObjIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .ValorAzul, .CanUse, .MinHitMag, .MaxHitMag, .MinDefMag, .MaxDefMag)
        End With

        'End If
    Next I
    
    frmComerciar.cantidad.Text = "1"
    frmComerciar.lblGld.Caption = PonerPuntos(UserGLD)
    
    Call Invalidate(frmComerciar.picInvNpc.hWnd)
    Call Invalidate(frmComerciar.picInvUser.hWnd)
    
    'Set state and show form
    Comerciando = True
    MirandoComerciar = True
    
    #If ModoBig = 1 Then
        dockForm frmComerciar.hWnd, FrmMain.PicMenu, True
    #Else
        frmComerciar.Show , FrmMain
    #End If

End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I          As Long

    'Dim BankGold   As Long

    'Dim BankEldhir As Long
    
    UserBankGold = Reader.ReadInt
    UserBankEldhir = Reader.ReadInt
    SelectedBank = Reader.ReadInt
    
    Set InvBanco(0) = New clsGrapchicalInventory
    Set InvBanco(1) = New clsGrapchicalInventory
    
    Call InvBanco(0).Initialize(frmBancoObj.PicBancoInv, 28, MAX_BANCOINVENTORY_SLOTS, eCaption.Boveda_Npc, , , , , , , , True, , True, , True)
    Call InvBanco(1).Initialize(frmBancoObj.picInv, 28, Inventario.MaxObjs, eCaption.Boveda_User, , , , , , , , , , True)
    
    For I = 1 To Inventario.MaxObjs

        With Inventario
            Call InvBanco(1).SetItem(I, .ObjIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .ValorAzul(I), .CanUse(I), .MinHitMag(I), .MaxHitMag(I), .MinDefMag(I), .MaxDefMag(I))

        End With

    Next I
    
    For I = 1 To MAX_BANCOINVENTORY_SLOTS

        With UserBancoInventory(I)
            Call InvBanco(0).SetItem(I, .ObjIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .ValorAzul, .CanUse, .MinHitMag, .MaxHitMag, .MinDefMag, .MaxDefMag)

        End With

    Next I
    
    'Set state and show form
    Comerciando = True
    
    'Call Invalidate(frmBancoObj.PicBancoInv.hWnd)
    'Call Invalidate(frmBancoObj.picInv.hWnd)
    
    'Call InvBanco(0).DrawInventory
    'Call InvBanco(1).DrawInventory
    
    MirandoBanco = True
    
    #If ModoBig = 1 Then
        dockForm frmBancoObj.hWnd, FrmMain.PicMenu, True
        
    #Else
        frmBancoObj.Show , FrmMain
    #End If
    
   ' lblGld.Caption = IIf(UserBankGold > 0, Format$(UserBankGold, "##,##"), "0")
'    lblDSP.Caption = IIf(UserBankEldhir > 0, Format$(UserBankEldhir, "##,##"), "0")
    
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I As Long
    
    TradingUserName = Reader.ReadString8

    ' Comercio entre usuarios
    Set InvComUsu = New clsGrapchicalInventory
    Set InvOfferComUsu(0) = New clsGrapchicalInventory
    Set InvOfferComUsu(1) = New clsGrapchicalInventory
    Set InvOroComUsu(0) = New clsGrapchicalInventory
    Set InvOroComUsu(1) = New clsGrapchicalInventory
    Set InvOroComUsu(2) = New clsGrapchicalInventory
    
    Set InvEldhirComUsu(0) = New clsGrapchicalInventory
    Set InvEldhirComUsu(1) = New clsGrapchicalInventory
    Set InvEldhirComUsu(2) = New clsGrapchicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(frmComerciarUsu.picInvComercio, Inventario.MaxObjs, Inventario.MaxObjs, eCaption.cInvComUsu)
    Call InvOfferComUsu(0).Initialize(frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS, INV_OFFER_SLOTS, eCaption.cInvOfferComUsu1)
    Call InvOfferComUsu(1).Initialize(frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS, INV_OFFER_SLOTS, eCaption.cInvOfferComUsu2)
    Call InvOroComUsu(0).Initialize(frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, INV_GOLD_SLOTS, eCaption.cInvOroComUsu1, TilePixelWidth, TilePixelHeight, TilePixelWidth, , , , True)
    Call InvOroComUsu(1).Initialize(frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, INV_GOLD_SLOTS, eCaption.cInvOroComUsu2, TilePixelWidth, TilePixelHeight, TilePixelWidth, , , , True)
    Call InvOroComUsu(2).Initialize(frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, INV_GOLD_SLOTS, eCaption.cInvOroComUsu3, TilePixelWidth, TilePixelHeight, TilePixelWidth, , , , True)
    Call InvEldhirComUsu(0).Initialize(frmComerciarUsu.picInvEldhirProp, INV_GOLD_SLOTS, INV_GOLD_SLOTS, eCaption.cInvEldhirComUsu1, TilePixelWidth, TilePixelHeight, TilePixelWidth, , , , True)
    Call InvEldhirComUsu(1).Initialize(frmComerciarUsu.picInvEldhirOfertaProp, INV_GOLD_SLOTS, INV_GOLD_SLOTS, eCaption.cInvEldhirComUsu2, TilePixelWidth, TilePixelHeight, TilePixelWidth, , , , True)
    Call InvEldhirComUsu(2).Initialize(frmComerciarUsu.picInvEldhirOfertaOtro, INV_GOLD_SLOTS, INV_GOLD_SLOTS, eCaption.cInvEldhirComUsu3, TilePixelWidth, TilePixelHeight, TilePixelWidth, , , , True)
    
    'Fill user inventory
    For I = 1 To MAX_INVENTORY_SLOTS

        If Inventario.ObjIndex(I) <> 0 Then

            With Inventario
                Call InvComUsu.SetItem(I, .ObjIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .ValorAzul(I), .CanUse(I), .MinHitMag(I), .MaxHitMag(I), .MinDefMag(I), .MaxDefMag(I))
            End With

        End If

    Next I

    ' Inventarios de oro
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro", 0, True, 0, 0, 0, 0)
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro", 0, True, 0, 0, 0, 0)
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro", 0, True, 0, 0, 0, 0)
    
    Call InvEldhirComUsu(0).SetItem(1, ELDHIR_INDEX, UserDSP, 0, ELDHIR_GRH, 0, 0, 0, 0, 0, 0, "Eldhir", 0, True, 0, 0, 0, 0)
    Call InvEldhirComUsu(1).SetItem(1, ELDHIR_INDEX, 0, 0, ELDHIR_GRH, 0, 0, 0, 0, 0, 0, "Eldhir", 0, True, 0, 0, 0, 0)
    Call InvEldhirComUsu(2).SetItem(1, ELDHIR_INDEX, 0, 0, ELDHIR_GRH, 0, 0, 0, 0, 0, 0, "Eldhir", 0, True, 0, 0, 0, 0)
    
    frmComerciarUsu.Form_LoadDetails
        
    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, FrmMain)
    
    'InvComUsu.DrawInventory
    'InvOfferComUsu(0).DrawInventory
   ' InvOfferComUsu(1).DrawInventory
    'InvOroComUsu(0).DrawInventory
   ' InvOroComUsu(1).DrawInventory
   ' InvOroComUsu(2).DrawInventory
   ' InvEldhirComUsu(0).DrawInventory
   ' InvEldhirComUsu(1).DrawInventory
   ' InvEldhirComUsu(2).DrawInventory
    Exit Sub

    MsgBox ("Error comerciar con usuarios")
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Destroy the form and reset the state
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvEldhirComUsu(0) = Nothing
    Set InvEldhirComUsu(1) = Nothing
    Set InvEldhirComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    Unload frmComerciarUsu
    Comerciando = False
End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    
    With frmComerciarUsu
        ' Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSE
    End With
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim A As Integer
    
    'Get data and update form
    UserMinSTA = Reader.ReadInt()
    
    For A = FrmMain.lblEnergia.LBound To FrmMain.lblEnergia.UBound
        FrmMain.lblEnergia(A) = UserMinSTA & "/" & UserMaxSTA
    Next A
    
    FrmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * SHAPE_LONGITUD)
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'Get data and update form
    UserMinMAN = Reader.ReadInt()
    Reader.ReadInt16
    Dim A As Integer
    
    For A = FrmMain.lblMana.LBound To FrmMain.lblMana.UBound
        FrmMain.lblMana(A) = UserMinMAN & "/" & UserMaxMAN
    Next A
    
    If UserMaxMAN > 0 Then
        FrmMain.MANShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * SHAPE_LONGITUD)
    Else
        FrmMain.MANShp.Width = 0
    End If

End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'Get data and update form
    UserMinHP = Reader.ReadInt()
    Reader.ReadInt16
    
    Dim A As Integer
    
    For A = FrmMain.lblVida.LBound To FrmMain.lblVida.UBound
        FrmMain.lblVida(A).Caption = UserMinHP & "/" & UserMaxHP
    Next A
    

    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1

        If FrmMain.TrainingMacro Then Call FrmMain.DesactivarMacroHechizos
        If FrmMain.MacroTrabajo Then Call FrmMain.DesactivarMacroTrabajo
        
        FrmMain.Hpshp.Width = 0
    Else
        FrmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * SHAPE_LONGITUD)
        UserEstado = 0
    End If

End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 09/21/10
    'Last Modified By: C4b3z0n
    '- 08/14/07: Tavo - Added GldLbl color variation depending on User Gold and Level
    '- 09/21/10: C4b3z0n - Modified color change of gold ONLY if the player's level is greater than 12 (NOT newbie).
    '***************************************************
    
    'Get data and update form
    UserGLD = Reader.ReadInt()
    
    Dim A As Long
    
    For A = FrmMain.GldLbl.LBound To FrmMain.GldLbl.UBound
        FrmMain.GldLbl(A).Caption = PonerPuntos(UserGLD)
    Next A
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateDsp()

    '***************************************************
    'Autor: WAICON
    'Last Modification: 06/05/2019
    '
    '***************************************************
    
    'Get data and update form
    UserDSP = Reader.ReadInt32()
    Account.Eldhir = Reader.ReadInt32
    
    FrmMain.lblEldhir.Caption = UserDSP
    
    If MirandoShop Then
        FrmShop.lblCantDSP.Caption = Account.Eldhir
    End If
End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    
    Dim Gld As Long
    Dim Eldhir As Long
    
    Gld = Reader.ReadInt
    Eldhir = Reader.ReadInt
    
    frmBancoObj.lblGld.Caption = IIf(Gld > 0, Format$(Gld, "##,##"), "0")
    frmBancoObj.lblDsp.Caption = IIf(Eldhir > 0, Format$(Eldhir, "##,##"), "0")
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'Get data and update form
    UserExp = Reader.ReadInt32()
    
    Call Render_Exp(True)
   
End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '***************************************************
    FrmMain.tmrBlink.Enabled = False
    
    'Get data and update form
    UserFuerza = Reader.ReadInt
    UserAgilidad = Reader.ReadInt
    
    
    Dim A As Long
    
    For A = FrmMain.lblFuerza.LBound To FrmMain.lblFuerza.UBound
        FrmMain.lblFuerza(A).Caption = UserFuerza
        FrmMain.lblAgilidad(A).Caption = UserAgilidad
        FrmMain.lblFuerza(A).ForeColor = getStrenghtColor(UserFuerza)
        FrmMain.lblAgilidad(A).ForeColor = getDexterityColor(UserAgilidad)
    Next A
    
    
    
    'GlobalCounters.StrenghtAndDextery = 5
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '***************************************************
    FrmMain.tmrBlink.Enabled = False
    
    Dim A As Long
    
    'Get data and update form
    UserFuerza = Reader.ReadInt
    For A = FrmMain.lblFuerza.LBound To FrmMain.lblFuerza.UBound
        FrmMain.lblFuerza(A).Caption = UserFuerza
        FrmMain.lblAgilidad(A).Caption = UserAgilidad
        FrmMain.lblFuerza(A).ForeColor = getStrenghtColor(UserFuerza)
        FrmMain.lblAgilidad(A).ForeColor = getDexterityColor(UserAgilidad)
    Next A
    
    
    'GlobalCounters.StrenghtAndDextery = 5
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '***************************************************
    
    'Get data and update form
    FrmMain.tmrBlink.Enabled = False
    UserAgilidad = Reader.ReadInt
    
    Dim A As Long
    
    For A = FrmMain.lblFuerza.LBound To FrmMain.lblFuerza.UBound
        FrmMain.lblFuerza(A).Caption = UserFuerza
        FrmMain.lblAgilidad(A).Caption = UserAgilidad
        FrmMain.lblFuerza(A).ForeColor = getStrenghtColor(UserFuerza)
        FrmMain.lblAgilidad(A).ForeColor = getDexterityColor(UserAgilidad)
    Next A
    
     
    'GlobalCounters.StrenghtAndDextery = 5
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
        '<EhHeader>
        On Error GoTo HandleChangeMap_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     UserMap = Reader.ReadInt()
102     UserMapName = Reader.ReadString8()
    
    
        #If FullScreen = 0 Then
104     Map_TimeRender = 2000
106     FrmMain.tMapName.Enabled = True
        #End If
        
108     If FileExist(App.path & Maps_FilePath & "Mapa" & UserMap & ".map", vbNormal) Then
        
            Dim A As Long
        
110         For A = 1 To MAX_GUILD_MEMBER
112             MiniMap_Friends(A).X = 0
114             MiniMap_Friends(A).Y = 0
116         Next A
        
118         Call SwitchMap(UserMap)
        
120         If UserCharIndex > 0 Then
122             g_Last_OffsetX = CharList(UserCharIndex).MoveOffsetX
124             g_Last_OffsetY = CharList(UserCharIndex).MoveOffsetY
            End If
        
126         Draw_MiniMap
        
        Else
            'no encontramos el mapa en el hd
128         MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
130         Call CloseClient
        End If

        '<EhFooter>
        Exit Sub

HandleChangeMap_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Protocol.HandleChangeMap " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
        '<EhHeader>
        On Error GoTo HandlePosUpdate_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        Dim TempPosX As Byte, TempPosY As Byte
    
100     TempPosX = Reader.ReadInt()
102     TempPosY = Reader.ReadInt()
    
        'Remove char from old position
104     If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
106         MapData(UserPos.X, UserPos.Y).CharIndex = 0
        
108         Call g_Swarm.Remove(5, UserCharIndex, 0, 0, 0, 0)
        
        End If
    
        'Set new pos
110     UserPos.X = TempPosX
112     UserPos.Y = TempPosY
    
        'Set char
114     MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
116     CharList(UserCharIndex).Pos = UserPos
    
        Dim RangeX As Single, RangeY As Single

118     Call GetCharacterDimension(UserCharIndex, RangeX, RangeY)
        
120     Call g_Swarm.Insert(5, UserCharIndex, UserPos.X, UserPos.Y, RangeX, RangeY)
    
122     Call Audio.UpdateSource(CharList(UserCharIndex).SoundSource, UserPos.X, UserPos.Y)
    
        'Are we under a roof?
124     bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

126     Draw_MiniMap
        '<EhFooter>
        Exit Sub

HandlePosUpdate_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Protocol.HandlePosUpdate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Select Case Reader.ReadInt()

        Case bCabeza
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)

        Case bBrazoDerecho
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)

        Case bPiernaDerecha
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(Reader.ReadInt()) & "!!", 255, 0, 0, True, False, True)

        Case bTorso
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(Reader.ReadInt() & "!!"), 255, 0, 0, True, False, True)
    End Select

End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(Reader.ReadInt()) & MENSAJE_2, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & CharList(Reader.ReadInt()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim attacker As String
    
    attacker = CharList(Reader.ReadInt()).Nombre
    
    Select Case Reader.ReadInt

        Case bCabeza
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bBrazoDerecho
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bPiernaDerecha
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bTorso
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select

End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Victim As String
    
    Victim = CharList(Reader.ReadInt()).Nombre
    
    Select Case Reader.ReadInt

        Case bCabeza
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bBrazoDerecho
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bPiernaDerecha
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)

        Case bTorso
            Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(Reader.ReadInt() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select

End Sub

Public Function MAC_GET() As String

    Dim colNetAdapters, objWMIService As Object, objItem As Object

    Dim strComputer As String

    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    
    For Each objItem In colNetAdapters

        MAC_GET = objItem.MACAddress
    Next
    
    Exit Function

End Function

Public Function HD_GET(ByVal s_Drive As String) As Long

    Dim o_Fso   As Scripting.FileSystemObject

    Dim o_Drive As Drive
            
    ' Creamos un nuevo objeto de tipo Scripting FileSystemObject
    Set o_Fso = New Scripting.FileSystemObject
            
    ' Si el Drive no es un vbnullstring
    If s_Drive <> "" Then
        ' Recuperamos el Drive para poder acceder _
          en las siguientes lineas
        Set o_Drive = o_Fso.GetDrive(s_Drive)
    End If
            
    With o_Drive
                
        ' Si está disponible
        If .IsReady Then
            HD_GET = Not .SerialNumber
        Else
            HD_GET = -1
        End If

    End With
            
    ' Eliminamos los objetos instanciados
    Set o_Drive = Nothing
    Set o_Fso = Nothing
            
    Exit Function

End Function

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim chat      As String

    Dim CharIndex As Integer

    Dim r         As Byte

    Dim g         As Byte

    Dim b         As Byte
    
    chat = Reader.ReadString16()
    CharIndex = Reader.ReadInt()
    
    r = Reader.ReadInt()
    g = Reader.ReadInt()
    b = Reader.ReadInt()
    
    'If charlist(CharIndex).IsNpc Then
    If LastCharIndex <> 0 Then
        Dialogos.RemoveDialog (LastCharIndex)
    End If

    LastCharIndex = CharIndex
    ' End If
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If CharList(CharIndex).Active Then Call Dialogos.CreateDialog(Trim$(chat), CharIndex, r, g, b)
    
End Sub

''
' Handles the ChatPersonalizado message.

Private Sub HandleChatPersonalizado()

    '***************************************************
    'Author: Juan Dalmasso (CHOTS)
    'Last Modification: 11/06/2011
    '***************************************************
    
    Dim chat      As String

    Dim CharIndex As Integer

    Dim Tipo      As Byte
    
    chat = Reader.ReadString16()
    CharIndex = Reader.ReadInt()
    
    Tipo = Reader.ReadInt()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If CharList(CharIndex).Active Then
        If esGM(CharIndex) Then
            Call Dialogos.CreateDialog(Trim$(chat), CharIndex, 240, 215, 25, Tipo)
            
        Else
            Call Dialogos.CreateDialog(Trim$(chat), CharIndex, ColoresDialogos(Tipo).r, ColoresDialogos(Tipo).g, ColoresDialogos(Tipo).b, Tipo)
        End If
        
    End If
    
End Sub

Private Sub HandleConsoleMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/05/11
    'D'Artagnan: Agrego la división de consolas
    '***************************************************
    
    Dim chat        As String

    Dim FontIndex   As Integer

    Dim MessageType As eMessageType

    Dim str         As String

    Dim r           As Byte

    Dim g           As Byte

    Dim b           As Byte
    
    Dim TempGuild As String
    
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt()
    MessageType = Reader.ReadInt()
    
    
    
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)
        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            g = 255
        Else
            g = Val(str)
        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)
        End If
        
        Call AddtoRichTextBox(FrmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        
    Else
    
        If InStr(1, chat, "[CLANES]") Then
            TempGuild = Replace(chat, "[CLANES]", vbNullString)
            Call DialogosClanes.PushBackText(ReadField(1, TempGuild, 126))
        End If
    
        'If UCase$(Left$(chat, 11)) <> "[SEGURIDAD]" Then
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(FrmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
            
            #If FullScreen = 1 Then
                If MessageType = eMessageType.cEvents_General Then
                    Call AddtoRichTextBox(FrmMain.ConsoleEvents, chat, .red, .green, .blue, .bold, .italic)
                    Call AddtoRichTextBox(FrmMain.ConsoleEvents, vbNullString, .red, .green, .blue, .bold, .italic)
                    
                ElseIf MessageType = eMessageType.cEvents_Curso Then
                    Call AddtoRichTextBox(FrmMain.ConsoleCurso, chat, .red, .green, .blue, .bold, .italic)
                End If
                
                
            #End If
        End With

        'End If
        
        ' Para no perder el foco cuando chatea por party
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            'If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If

    '    Call checkText(chat)
    
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    '
    '***************************************************
    
    Dim chat      As String

    Dim FontIndex As Integer

    Dim str       As String

    Dim r         As Byte

    Dim g         As Byte

    Dim b         As Byte
    
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)
        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            g = 255
        Else
            g = Val(str)
        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)
        End If
            
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .red, .green, .blue, .bold, .italic)
        End With

    End If
    
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    frmMensaje.msg.Caption = Reader.ReadString8()
    frmMensaje.Show
    
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserIndex = Reader.ReadInt()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserCharIndex = Reader.ReadInt()
    UserPos = CharList(UserCharIndex).Pos
    'Draw_MiniMap
End Sub

Private Sub HandleCharacterChangeHeading()
    
    Dim CharIndex As Integer

    CharIndex = Reader.ReadInt
    
    CharList(CharIndex).Heading = Reader.ReadInt
    
    If CharList(CharIndex).Pos.X > 0 Then
    ' Sistema de Escaleras 65562 65563
    If MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).Graphic(2).GrhIndex = 65562 Or _
        MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).Graphic(2).GrhIndex = 65563 Then
        
        CharList(CharIndex).Heading = NORTH
    End If
    End If
    
    Dim RangeX As Single, RangeY As Single

    Call GetCharacterDimension(CharIndex, RangeX, RangeY)
        
    Call g_Swarm.Resize(CharIndex, RangeX, RangeY)
    
    Call RefreshAllChars
    
      
    
    With CharList(CharIndex)
                If Not .Moving Then

                'Start animations
                If .Body.Walk(.Heading).started = 0 Then
                    .Body.Walk(.Heading).started = FrameTime
                    .Arma.WeaponWalk(.Heading).started = FrameTime
                    .Escudo.ShieldWalk(.Heading).started = FrameTime

                    .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                    .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS

                End If
            
                .MovArmaEscudo = False
                .Moving = True

            End If
            
    End With
    
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex                 As Integer

    Dim Body                      As Integer

    Dim BodyAttack                As Integer

    Dim Head                      As Integer

    Dim Heading                   As E_Heading

    Dim X                         As Byte

    Dim Y                         As Byte

    Dim Weapon                    As Integer

    Dim Shield                    As Integer

    Dim helmet                    As Integer

    Dim privs                     As Integer

    Dim NickColor                 As Byte

    Dim AuraIndex(1 To MAX_AURAS) As Byte

    Dim ValidInvi                 As Boolean
    
    Dim NpcIndex                  As Integer
    
    Dim Pos As Integer
    
    CharIndex = Reader.ReadInt()
    Body = Reader.ReadInt()
    BodyAttack = Reader.ReadInt()
    Head = Reader.ReadInt()
    Heading = Reader.ReadInt()
    X = Reader.ReadInt()
    Y = Reader.ReadInt()
    Weapon = Reader.ReadInt()
    Shield = Reader.ReadInt()
    helmet = Reader.ReadInt()
    
    With CharList(CharIndex)
        .FxIndex = Reader.ReadInt()
        Reader.ReadInt
        If .FxIndex > 0 Then Call InitGrh(.fX, FxData(.FxIndex).Animacion)
        
        .Nombre = Reader.ReadString8()
        
        Pos = getTagPosition(.Nombre)
        .GuildName = mid$(.Nombre, Pos)
        
        
        NickColor = Reader.ReadInt()
        
        If (NickColor And eNickColor.ieCriminal) <> 0 Or (NickColor And eNickColor.ieCAOS) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0

        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        .ColorNick = NickColor
       
        .Muerto = (Body = iCuerpoMuerto) Or (Body = iCuerpoMuerto_Legion) Or (Body = FRAGATA_FANTASMAL)
    
        ' If CharIndex = UserCharIndex Then
        ' If .Muerto Then
        'Set_engineBaseSpeed (0.024)
        ' Else
        'Set_engineBaseSpeed (0.018)
        'End If
        '   End If
        
        privs = Reader.ReadInt()
        
        Dim A As Long
        
        For A = 1 To MAX_AURAS
            AuraIndex(A) = Reader.ReadInt()
        Next A
        
        .NpcIndex = Reader.ReadInt16
        
        Dim Flags As Byte

        Flags = Reader.ReadInt8
        
        .Idle = Flags And &O1
        .Navegando = Flags And &O2
        .Speeding = Reader.ReadReal32
        
        If privs <> 0 Then

            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                
                privs = privs Xor PlayerType.ChaosCouncil

            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil

            End If

            'If privs = 0 Then privs = 1
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .Priv = Log(privs) / Log(2)
            
        Else
            .Priv = 0

        End If

        Call MakeChar(CharIndex, Body, BodyAttack, Head, Heading, X, Y, Weapon, Shield, helmet, AuraIndex, .NpcIndex)
        
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).started = FrameTime

        End If

    End With
    
    Call RefreshAllChars
    
End Sub

Private Sub HandleCharacterChangeNick()

    '***************************************************
    'Author: Budi
    'Last Modification: 07/23/09
    '
    '***************************************************

    Dim CharIndex As Integer
    Dim Pos As Integer
    
    CharIndex = Reader.ReadInt
    CharList(CharIndex).Nombre = Reader.ReadString8
    Pos = getTagPosition(CharList(CharIndex).Nombre)
    CharList(CharIndex).GuildName = mid$(CharList(CharIndex).Nombre, Pos)
        
    If CharIndex = UserCharIndex Then
        FrmMain.Label8(0).Caption = CharList(CharIndex).Nombre
    End If
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
        '<EhHeader>
        On Error GoTo HandleCharacterRemove_Err
        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        Dim CharIndex As Integer
    
100     CharIndex = Reader.ReadInt()
    
102     Call EraseChar(CharIndex)
104     Call RefreshAllChars
        '<EhFooter>
        Exit Sub

HandleCharacterRemove_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Protocol.HandleCharacterRemove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex As Integer

    Dim X         As Byte

    Dim Y         As Byte
    
    CharIndex = Reader.ReadInt()
    X = Reader.ReadInt()
    Y = Reader.ReadInt()
    
    With CharList(CharIndex)
        
        ' Play steps sounds if the user is not an admin of any kind
        If .Priv <> 1 And .Priv <> 2 And .Priv <> 3 And .Priv <> 25 Then
            Call DoPasosFx(CharIndex)

        End If

    End With
    
    Call MoveCharbyPos(CharIndex, X, Y)
    
    Call RefreshAllChars

End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    Dim Direccion As Byte
    
    Direccion = Reader.ReadInt()

    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)
    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 21/09/2010 - C4b3z0n
    '25/08/2009: ZaMa - Changed a variable used incorrectly.
    '21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
    '***************************************************
    
    Dim CharIndex    As Integer

    Dim tempint      As Integer

    Dim tempbyt      As Byte

    Dim headIndex    As Integer
    
    Dim ModoStreamer As Boolean
    
    CharIndex = Reader.ReadInt()
    
    With CharList(CharIndex)
        tempint = Reader.ReadInt()
        
        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(tempint)
            .iBody = tempint

        End If
        
        tempint = Reader.ReadInt()
        
        If tempint < LBound(BodyDataAttack()) Or tempint > UBound(BodyDataAttack()) Then
            .BodyAttack = BodyDataAttack(0)
            .iBodyAttack = 0
        Else
            .BodyAttack = BodyDataAttack(tempint)
            .iBodyAttack = tempint

        End If

        headIndex = Reader.ReadInt()
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex

        End If
        
        .Muerto = (.iBody = iCuerpoMuerto) Or (.iBody = iCuerpoMuerto_Legion) Or (.iBody = FRAGATA_FANTASMAL)
        
        ' If CharIndex = UserCharIndex Then
        ' If .Muerto Then
        '  Set_engineBaseSpeed (0.024)
        'Else
        ' Set_engineBaseSpeed (0.018)
        ' End If
        ' End If
        
        .Heading = Reader.ReadInt()
        
        tempint = Reader.ReadInt()

        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)
        
        tempint = Reader.ReadInt()

        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
        tempint = Reader.ReadInt()

        If tempint <> 0 Then .Casco = CascoAnimData(tempint)

        If tempint = 54 Then
            .OffsetY = -15
        Else
            .OffsetY = 0

        End If
        
        .FxIndex = Reader.ReadInt
        
        Reader.ReadInt 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)
        End If
        
        'Call SetCharacterFx(charindex, Reader.ReadInt(), Reader.ReadInt())
        
        Dim A As Long

        For A = 1 To MAX_AURAS
            tempbyt = Reader.ReadInt()

            If tempbyt < LBound(AuraAnimData()) Or tempbyt > UBound(AuraAnimData()) Then
                .Aura(A) = AuraAnimData(0)
            Else
                .Aura(A) = AuraAnimData(tempbyt)

            End If
   
        Next A
        
        ModoStreamer = Reader.ReadBool
        .Streamer = ModoStreamer

        Dim Flags As Byte
        
        Flags = Reader.ReadInt8()
        
        .Idle = Flags And &O1
        .Navegando = Flags And &O2

        Dim RangeX As Single, RangeY As Single

        Call GetCharacterDimension(CharIndex, RangeX, RangeY)
        
        Call g_Swarm.Resize(CharIndex, RangeX, RangeY)
        
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).started = FrameTime

        End If

    End With
    
    Call RefreshAllChars

End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim X        As Byte

    Dim Y        As Byte

    Dim GrhIndex As Long
    
    Dim Name As String
    
    Dim ObjIndex As Integer
    
    Dim Amount As Long
    
    Dim Sound As Integer
    
    
    X = Reader.ReadInt()
    Y = Reader.ReadInt()
    GrhIndex = Reader.ReadInt()
    ObjIndex = Reader.ReadInt16
    Name = ObjData(ObjIndex).Name
    Amount = Reader.ReadInt
    Sound = Reader.ReadInt16
    
    With MapData(X, Y)
        .ObjName = Name
        .OBJInfo.Amount = Amount
        .OBJInfo.ObjIndex = ObjIndex
        
        If (.ObjGrh.GrhIndex <> 0) Then
            
            With GrhData(.ObjGrh.GrhIndex)
                Call g_Swarm.Remove(4, -1, X, Y, .TileWidth, .TileHeight)
            End With

        End If
            
80      .ObjGrh.GrhIndex = GrhIndex
            
        ' RTREE
        If (.ObjGrh.GrhIndex <> 0) Then
            
            
            With GrhData(.ObjGrh.GrhIndex)
                Call g_Swarm.Insert(4, -1, X, Y, .TileWidth, .TileHeight)
            End With
            
            
        End If
       
       
       ' Hacer sistema de Sonido en Tiles
       'Or (GrhIndex = GrhFogata2)
       
       
       If (Sound <> 0) Then
            .OBJInfo.SoundSource = Audio.CreateSource(X, Y)
                   
            Call Audio.PlayEffect(CStr(Sound) & ".wav", .OBJInfo.SoundSource, True)
       End If
        
90

Call InitGrh(.ObjGrh, .ObjGrh.GrhIndex)

    End With

End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim X As Byte

    Dim Y As Byte
    
    X = Reader.ReadInt()
    Y = Reader.ReadInt()
    
    Call Audio.DeleteSource(MapData(X, Y).OBJInfo.SoundSource, True)
    
    With GrhData(MapData(X, Y).ObjGrh.GrhIndex)
        Call g_Swarm.Remove(4, -1, X, Y, .TileWidth, .TileHeight)
    End With
          
    MapData(X, Y).ObjGrh.GrhIndex = 0
    MapData(X, Y).OBJInfo.ObjIndex = 0
    MapData(X, Y).OBJInfo.Amount = 0
    MapData(X, Y).ObjName = vbNullString
    
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim X As Byte

    Dim Y As Byte
    
    X = Reader.ReadInt()
    Y = Reader.ReadInt()
    
    If Reader.ReadBool() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If

End Sub

''
' Handles the PlayMusic message.

Private Sub HandlePlayMusic()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Music As Integer
    Music = Reader.ReadInt()
    
    Call Audio.PlayMusic(CStr(Music) + ".mp3")

End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
        
    Dim Wave    As Integer

    Dim SrcX    As Byte

    Dim SrcY    As Byte

    Dim Entity  As Long

    Dim Repeat  As Boolean

    Dim MapOnly As Boolean
    
    Wave = Reader.ReadInt()
    SrcX = Reader.ReadInt()
    SrcY = Reader.ReadInt()
    Entity = Reader.ReadInt()
    Repeat = Reader.ReadBool()
    MapOnly = Reader.ReadBool()
    
    
    If Wave = 0 Then
         With MapData(SrcX, SrcY)
            If .SoundSource > 0 Then
                Call Audio.DeleteSource(.SoundSource, True)
            End If
        End With
    Else
        If MapOnly Then
    
            With MapData(SrcX, SrcY)
            
                If .SoundSource > 0 Then
                    Call Audio.DeleteSource(.SoundSource, True)
                End If
            
                .SoundSource = Audio.CreateSource(SrcX, SrcY)
                Call Audio.PlayEffect(CStr(Wave) & ".wav", .SoundSource, Repeat)
            
            End With
          
        Else
            
            If SrcX = 0 And SrcY = 0 Then
                SrcX = UserPos.X
                SrcY = UserPos.Y
            End If
            
            If (Entity > 0) Then
                Call Audio.PlayEffect(Wave & ".wav", CharList(Entity).SoundSource, Repeat)
            Else
                Call Audio.PlayEffectAt(Wave & ".wav", SrcX, SrcY)
    
            End If
    
        End If
    End If
    
    

End Sub

Private Sub HandleStopWave()
    
    Dim X As Byte, Y As Byte, Inmediatily As Boolean
    
    X = Reader.ReadInt8
    Y = Reader.ReadInt8
    Inmediatily = Reader.ReadBool
    
    With MapData(X, Y)
        
        If .SoundSource > 0 Then
            Call Audio.DeleteSource(.SoundSource, Inmediatily)

        End If

    End With

End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    pausa = Not pausa
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex    As Integer

    Dim fX           As Integer

    Dim Loops        As Integer
    
    Dim IsMeditation As Boolean
    
    CharIndex = Reader.ReadInt()
    fX = Reader.ReadInt()
    Loops = Reader.ReadInt()
    IsMeditation = Reader.ReadBool
    
    If fX = 0 Then
        CharList(CharIndex).fX.AnimacionContador = 29
        Exit Sub

    End If
    
    Call SetCharacterFx(CharIndex, fX, Loops)

End Sub

Private Sub HandleCreateFXMap()
    
    Dim X As Byte
    Dim Y As Byte
    Dim fX As Integer
    Dim FxLoops As Integer
    
    X = Reader.ReadInt
    Y = Reader.ReadInt
    fX = Reader.ReadInt
    FxLoops = Reader.ReadInt
    
    Dim GrhIndex As Long
    Dim Animacion As Long
    
    If fX = 0 Then
        Call g_Swarm.Remove(6, -1, X, Y, 2, 2)
        Exit Sub
    End If
    
    Animacion = FxData(fX).Animacion
    GrhIndex = GrhData(Animacion).Frames(1)
    
    With GrhData(Animacion)
        
        If MapData(X, Y).FxIndex > 0 Then
            Call g_Swarm.Remove(6, -1, X, Y, 2, 2)
        End If
    
        Call g_Swarm.Insert(6, -1, X, Y, 2, 2)
        
        Call SetCharacterFxMap(X, Y, fX, FxLoops)
    End With
    
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()

        '<EhHeader>
        On Error GoTo HandleUpdateUserStats_Err

        '</EhHeader>

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        Dim UserEstatus As Byte

100     UserMaxHP = Reader.ReadInt16()
102     UserMinHP = Reader.ReadInt16()
104     UserMaxMAN = Reader.ReadInt16()
106     UserMinMAN = Reader.ReadInt16()
108     UserMaxSTA = Reader.ReadInt16()
110     UserMinSTA = Reader.ReadInt16()
112     UserGLD = Reader.ReadInt32()
114     UserDSP = Reader.ReadInt32()
116     UserLvl = Reader.ReadInt8()
118     UserPasarNivel = Reader.ReadInt32()
120     UserExp = Reader.ReadInt32()
122     UserPoints = Reader.ReadInt32()
    
        ' Este es ultimo byte leido
124     UserEstatus = Reader.ReadInt8()
    
126     Select Case UserEstatus

            Case 0
128             FrmMain.lblStatus(0).visible = False

130         Case 1
132             FrmMain.lblStatus(0).Caption = "[AVENTURERO]"
134             FrmMain.lblStatus(0).ForeColor = RGB(90, 50, 20)
136             FrmMain.lblStatus(0).visible = True

138         Case 2
140             FrmMain.lblStatus(0).Caption = "[HEROE]"
142             FrmMain.lblStatus(0).ForeColor = RGB(170, 150, 150)
144             FrmMain.lblStatus(0).visible = True

146         Case 3
148             FrmMain.lblStatus(0).Caption = "[LEYENDA]"
150             FrmMain.lblStatus(0).ForeColor = vbYellow
152             FrmMain.lblStatus(0).visible = True
        
        End Select
    
        Dim A As Long
    
154     Call Render_Exp(True)

156     For A = FrmMain.GldLbl.LBound To FrmMain.GldLbl.UBound
158         FrmMain.GldLbl(A).Caption = PonerPuntos(UserGLD)
160     Next A

162     FrmMain.lblEldhir.Caption = UserDSP
    
        'Stats
164     For A = FrmMain.lblVida.LBound To FrmMain.lblVida.UBound
166         FrmMain.lblVida(A) = UserMinHP & "/" & UserMaxHP
168     Next A
    
170     For A = FrmMain.lblMana.LBound To FrmMain.lblMana.UBound
172         FrmMain.lblMana(A) = UserMinMAN & "/" & UserMaxMAN
174     Next A
    
176     For A = FrmMain.lblEnergia.LBound To FrmMain.lblEnergia.UBound
178         FrmMain.lblEnergia(A) = UserMinSTA & "/" & UserMaxSTA
180     Next A
    
        If UserMinSTA > 0 Then
182         FrmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * SHAPE_LONGITUD)
        Else
            FrmMain.STAShp.Width = 0

        End If
        
184     If UserMinHP > 0 Then
186         FrmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * SHAPE_LONGITUD)
        Else
188         FrmMain.Hpshp.Width = 0

        End If
    
190     If UserMaxMAN > 0 Then
192         FrmMain.MANShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * SHAPE_LONGITUD)
        Else
            FrmMain.MANShp.Width = 0

        End If
    
194     If UserMinHP = 0 Then
196         UserEstado = 1

198         If FrmMain.TrainingMacro Then Call FrmMain.DesactivarMacroHechizos
200         If FrmMain.MacroTrabajo Then Call FrmMain.DesactivarMacroTrabajo
        Else
202         UserEstado = 0

        End If
    
204     Call Mod_General.Render_Exp(True)

        '<EhFooter>
        Exit Sub

HandleUpdateUserStats_Err:
        LogError err.Description & vbCrLf & "in ARGENTUM.Protocol.HandleUpdateUserStats " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Slot      As Byte

    Dim ObjIndex  As Long

    Dim Name      As String

    Dim Amount    As Integer

    Dim Equipped  As Boolean

    Dim GrhIndex  As Long

    Dim ObjType   As Byte

    Dim MaxHit    As Integer

    Dim MinHit    As Integer

    Dim MaxDef    As Integer

    Dim MinDef    As Integer

    Dim Value     As Single

    Dim ValueAzul As Single

    Dim CanUse    As Boolean

    Dim MinHitMag As Integer

    Dim MaxHitMag As Integer

    Dim MinDefMag As Integer

    Dim MaxDefMag As Integer

    Dim A         As Long
    
    Dim Time As Long
    
    Time = FrameTime
    
    Dim Bronce    As Byte, Plata As Byte, Oro As Byte, Premium As Byte
    Slot = Reader.ReadInt()
    ObjIndex = Reader.ReadInt()
    Name = Reader.ReadString8()
    Amount = Reader.ReadInt()
    Equipped = Reader.ReadBool()
    GrhIndex = Reader.ReadInt()
    ObjType = Reader.ReadInt()
    MaxHit = Reader.ReadInt()
    MinHit = Reader.ReadInt()
    MaxDef = Reader.ReadInt()
    MinDef = Reader.ReadInt()
    Value = Reader.ReadReal32()
    ValueAzul = Reader.ReadReal32()
    CanUse = Reader.ReadBool()
        
    MinHitMag = Reader.ReadInt()
    MaxHitMag = Reader.ReadInt()
    MinDefMag = Reader.ReadInt()
    MaxDefMag = Reader.ReadInt()
    
    Bronce = Reader.ReadInt8()
    Plata = Reader.ReadInt8()
    Oro = Reader.ReadInt8()
    Premium = Reader.ReadInt8()
    
    If Equipped Then
        
        Select Case ObjType
            
            Case eOBJType.otMagic
                Dim Porc As String
                    Select Case ObjIndex
                    
                        Case LAUDMAGICO
                            Porc = "5%"
                        Case ANILLOMAGICO
                            Porc = "3%"
                    End Select
                    
                 FrmMain.lblMagic.Caption = Porc
                 UserMagicEqpSlot = Slot
                 
            Case eOBJType.otAnillo
                FrmMain.lblAnillo.Caption = MaxDefMag
                UserAnilloEqpSlot = Slot
                
            Case eOBJType.otWeapon
                FrmMain.lblWeapon = MaxHit
                UserWeaponEqpSlot = Slot
                If ObjData(ObjIndex).DamageMag > 0 Then
                    FrmMain.lblMagic.Caption = (ObjData(ObjIndex).DamageMag + 70) / 100 & "%"
                End If
                
            Case eOBJType.otarmadura
                FrmMain.lblarmor = MaxDef
                UserArmourEqpSlot = Slot
                
            Case eOBJType.otescudo

                FrmMain.lblShielder = MaxDef

                UserHelmEqpSlot = Slot

            Case eOBJType.otcasco

                FrmMain.lblhelm = MaxDef

                UserShieldEqpSlot = Slot
        End Select

    Else

        Select Case Slot
            Case UserMagicEqpSlot
              FrmMain.lblMagic.Caption = "0%" ' Esto tiene que venir cargado
            UserMagicEqpSlot = 0
                 
            Case UserAnilloEqpSlot
                FrmMain.lblAnillo.Caption = "0"
                UserAnilloEqpSlot = 0
                
            Case UserWeaponEqpSlot
                FrmMain.lblMagic.Caption = "0%"
                
                FrmMain.lblWeapon = "0"

                UserWeaponEqpSlot = 0

            Case UserArmourEqpSlot

                FrmMain.lblarmor = "0"

                UserArmourEqpSlot = 0

            Case UserHelmEqpSlot

                FrmMain.lblShielder = "0"

                UserHelmEqpSlot = 0

            Case UserShieldEqpSlot

                FrmMain.lblhelm = "0"

                UserShieldEqpSlot = 0
        End Select

    End If
    
    Call Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, _
       MaxDef, MinDef, Value, Name, ValueAzul, CanUse, MinHitMag, MaxHitMag, MinDefMag, MaxDefMag, Bronce, Plata, Oro, Premium)
    
    
 ' If Time - LastUpdateInv > 100 Then
        Inventario.DrawInventory
        'LastUpdateInv = Time

  '  End If
End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()

    '***************************************************
    'Author: Budi
    'Last Modification: 12/01/09
    '
    '***************************************************
    
    MaxInventorySlots = Reader.ReadInt
End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/03/10
    '
    '***************************************************
    Dim Slot   As Byte

    Dim Amount As Long
    
    Slot = Reader.ReadInt
    
    With InvOfferComUsu(0)
        Amount = .Amount(Slot)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.ObjIndex(Slot), Amount)
            
            ' Borro el item
            Call .SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, False, 0, 0, 0, 0)
        End If

    End With
    
    ' Si era el único ítem de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) And Not frmComerciarUsu.HasAnyItem(InvEldhirComUsu(1)) Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)
    End With

End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Slot As Byte

    Slot = Reader.ReadInt()
    
    With UserBancoInventory(Slot)
        .ObjIndex = Reader.ReadInt()
        .Name = Reader.ReadString8()
        .Amount = Reader.ReadInt()
        .GrhIndex = Reader.ReadInt()
        .ObjType = Reader.ReadInt()
        .MaxHit = Reader.ReadInt()
        .MinHit = Reader.ReadInt()
        .MaxDef = Reader.ReadInt()
        .MinDef = Reader.ReadInt
        .Valor = Reader.ReadInt()
        .ValorAzul = Reader.ReadInt()
        .CanUse = Reader.ReadBool()
            
        .MinHitMag = Reader.ReadInt()
        .MaxHitMag = Reader.ReadInt()
        .MinDefMag = Reader.ReadInt()
        .MaxDefMag = Reader.ReadInt()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(Slot, .ObjIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .ValorAzul, .CanUse, .MinHitMag, .MaxHitMag, .MinDefMag, .MaxDefMag)
        End If

    End With
    
End Sub

Private Sub HandleChangeBankSlot_Account()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Slot As Byte

    Slot = Reader.ReadInt()
    
    With UserBancoInventory(Slot)
        .ObjIndex = Reader.ReadInt()
        .Name = Reader.ReadString8()
        .Amount = Reader.ReadInt()
        .GrhIndex = Reader.ReadInt()
        .ObjType = Reader.ReadInt()
        .MaxHit = Reader.ReadInt()
        .MinHit = Reader.ReadInt()
        .MaxDef = Reader.ReadInt()
        .MinDef = Reader.ReadInt
        .Valor = Reader.ReadInt()
        .ValorAzul = Reader.ReadInt()
        .CanUse = Reader.ReadBool()
            
        .MinHitMag = Reader.ReadInt()
        .MaxHitMag = Reader.ReadInt()
        .MinDefMag = Reader.ReadInt()
        .MaxDefMag = Reader.ReadInt()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(Slot, .ObjIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .ValorAzul, .CanUse, .MinHitMag, .MaxHitMag, .MinDefMag, .MaxDefMag)
        End If

    End With
    
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Slot As Byte

    Slot = Reader.ReadInt()
    
    UserHechizos(Slot) = Reader.ReadInt()
    
    If Slot <= hlst.ListCount Then
        hlst.List(Slot - 1) = Reader.ReadString8()
    Else
        Call hlst.AddItem(Reader.ReadString8())
        hlst.Scroll = LastScroll
    End If
    
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim I As Long
    
    For I = 1 To NUMATRIBUTES
        UserAtributos(I) = Reader.ReadInt()
    Next I
    
End Sub



''
' Handles the RestOK message.

Private Sub HandleRestOK()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserDescansar = Not UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call MsgBox(Reader.ReadString8())
    
    'If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
        'If modNetwork.IsConnected Then
            'Call modNetwork.Disconnect
        'End If
    'End If
    
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserEstupido = True
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Slot As Byte

    Slot = Reader.ReadInt()
    
    With NPCInventory(Slot)
        .ObjIndex = Reader.ReadInt()
        
        .Name = Reader.ReadString8()
        .Amount = Reader.ReadInt()
        .Valor = Reader.ReadReal32()
        .GrhIndex = Reader.ReadInt()
        .ObjType = Reader.ReadInt()
        .MaxHit = Reader.ReadInt()
        .MinHit = Reader.ReadInt()
        .MaxDef = Reader.ReadInt()
        .MinDef = Reader.ReadInt()
        .ValorAzul = Reader.ReadReal32()
        .CanUse = Reader.ReadBool()
        
        .MinHitMag = Reader.ReadInt()
        .MaxHitMag = Reader.ReadInt()
        .MinDefMag = Reader.ReadInt()
        .MaxDefMag = Reader.ReadInt()
        
        .Animation = Reader.ReadInt()
    End With
    
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserMaxAGU = Reader.ReadInt()
    UserMinAGU = Reader.ReadInt()
    UserMaxHAM = Reader.ReadInt()
    UserMinHAM = Reader.ReadInt()
    
    Dim A As Integer
    
    For A = FrmMain.lblsed.LBound To FrmMain.lblsed.UBound
        FrmMain.Lblham(A) = UserMinHAM & "/" & UserMaxHAM
        FrmMain.lblsed(A) = UserMinAGU & "/" & UserMaxAGU
    Next A

    If UserMaxAGU > 0 Then
    FrmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * SHAPE_LONGITUD_MITAD)
    FrmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * SHAPE_LONGITUD_MITAD)
    Else
    FrmMain.AGUAsp.Width = 0
    FrmMain.COMIDAsp.Width = 0
    End If
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
    
    With UserEstadisticas
        .FragsCiu = Reader.ReadInt16()
        .FragsCri = Reader.ReadInt16()

        .Clase = Reader.ReadInt8()
        .Raza = Reader.ReadInt8()
        .Promedy = Reader.ReadInt32()
        
        .Elv = Reader.ReadInt8()
        .Exp = Reader.ReadInt32()
        .Elu = Reader.ReadInt32()
        
    End With
    
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    SkillPoints = Reader.ReadInt()

End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex As Integer
    
    CharIndex = Reader.ReadInt()
    CharList(CharIndex).Invisible = Reader.ReadBool()
    CharList(CharIndex).Intermitencia = Reader.ReadBool()
    
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleMeditateToggle_Err
    
    Dim CharIndex As Integer, fX As Integer
    Dim IUserMeditar As Boolean
    Dim X         As Integer, Y As Integer
    
    CharIndex = Reader.ReadInt16
    fX = Reader.ReadInt16
    X = Reader.ReadInt16
    Y = Reader.ReadInt16
    IUserMeditar = Reader.ReadBool()
    
    'If x + y > 0 Then
    'With CharList(charindex)
    'If .Invisible And charindex <> UserCharIndex Then
    ' If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
    '.Pos.x = x
    '  .Pos.y = y
    '  MapData(rrX(x), rrY(y)).charindex = charindex
    ' End If
    'End With
    ' End If
    
    If IUserMeditar Then
        If CharIndex = UserCharIndex Then
            UserMeditar = (fX <> 0)
            
            If UserMeditar Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("Comienzas a meditar.", .red, .green, .blue, .bold, .italic)
    
                End With
    
            Else
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("Has dejado de meditar.", .red, .green, .blue, .bold, .italic)
    
                End With
    
            End If
    
        End If
    End If
    
    With CharList(CharIndex)

        If fX <> 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)

        End If
        
        .FxIndex = fX
        .fX.Loops = -1
        .fX.AnimacionContador = 0

    End With
    
    Exit Sub

HandleMeditateToggle_Err:

End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Now the server send the percentage of progress of the skills.
    '***************************************************
    
    Dim I As Long
    Dim UserClase As Byte
    
    UserClase = Reader.ReadInt8
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = Reader.ReadInt8
        UserEstadisticas.Skills_Valid(I) = Skills_ValidateUser(UserClase, I)
        'Call Reader.ReadSafeArrayInt8(UserSkills)
    Next I
    
    #If ModoBig = 1 Then
        dockForm frmSkills3.hWnd, FrmMain.PicMenu, True
    #Else
        frmSkills3.Show vbModeless, FrmMain
    #End If
End Sub

Public Function Skills_ValidateUser(ByVal Clase As eClass, ByVal Skill As Byte) As Boolean

    ' Skills para TODOS
    If Skill = eSkill.Resistencia Or _
       Skill = eSkill.Tacticas Or _
       Skill = eSkill.Navegacion Then
        
        Skills_ValidateUser = True
        Exit Function
            
    End If

    ' Solo Trabajadores
        If Skill = eSkill.Mineria Or _
           Skill = eSkill.Talar Or _
           Skill = eSkill.Pesca Then
                
            Skills_ValidateUser = True
                    
            Exit Function
        End If
        
    ' Todas menos MAGO
    If Clase <> eClass.Mage Then
        If Skill = eSkill.Armas Or _
           Skill = eSkill.Apuñalar Then
            
            Skills_ValidateUser = True
            
            Exit Function
        End If
    End If
    
    ' Todas menos MAGO-DRUIDA
    If Clase <> eClass.Mage And _
       Clase <> eClass.Druid Then
        
        If Skill = eSkill.Defensa Then
            Skills_ValidateUser = True
            Exit Function
        End If
        
    End If

    ' Ladron
    If Clase = eClass.Thief Then
        
        If Skill = eSkill.Robar Then
            Skills_ValidateUser = True
            Exit Function
        End If
        
        If Skill = eSkill.Ocultarse Then
            Skills_ValidateUser = True
            Exit Function
        End If
        
        If Skill = eSkill.Proyectiles Then
            Skills_ValidateUser = True
            Exit Function
        End If
    End If
    
    ' Cazadores
    If Clase = eClass.Hunter Then
        If Skill = eSkill.Ocultarse Then
            Skills_ValidateUser = True
            Exit Function
        End If
        
        If Skill = eSkill.Proyectiles Then
            Skills_ValidateUser = True
            Exit Function
        End If
    End If
    
    ' Guerreros
    If Clase = eClass.Hunter Then
        If Skill = eSkill.Proyectiles Then
            Skills_ValidateUser = True
            Exit Function
        End If
    End If
    

End Function
''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserParalizado = Not UserParalizado
    
    If Not UserParalizado Then
        FrmMain.imgParalisis.visible = False
        FrmMain.lblParalisis.visible = False
        GlobalCounters.Paralized = 0
    End If
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    Call frmUserRequest.Show(vbModeless, FrmMain)
    
End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    If frmComerciar.visible Then

        Dim I As Long
        
        'Update user inventory
        For I = 1 To MAX_INVENTORY_SLOTS

            ' Agrego o quito un item en su totalidad
            If Inventario.ObjIndex(I) <> InvComUser.ObjIndex(I) Then

                With Inventario
                    Call InvComUser.SetItem(I, .ObjIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .ValorAzul(I), .CanUse(I), .MinHitMag(I), .MaxHitMag(I), .MinDefMag(I), .MaxDefMag(I))
                End With

                ' Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.Amount(I) <> InvComUser.Amount(I) Then
                Call InvComUser.ChangeSlotItemAmount(I, Inventario.Amount(I))
            End If

        Next I
        
        ' Fill Npc inventory
        For I = 1 To 20

            ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(I).ObjIndex <> InvComNpc.ObjIndex(I) Then

                With NPCInventory(I)
                    Call InvComNpc.SetItem(I, .ObjIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .ValorAzul, .CanUse, .MinHitMag, .MaxHitMag, .MinDefMag, .MaxDefMag)
                End With

                ' Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(I).Amount <> InvComNpc.Amount(I) Then
                Call InvComNpc.ChangeSlotItemAmount(I, NPCInventory(I).Amount)
            End If

        Next I
    
    End If

End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim I As Long
    
    If frmBancoObj.visible Then
        
        For I = 1 To Inventario.MaxObjs

            With Inventario
                Call InvBanco(1).SetItem(I, .ObjIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .ValorAzul(I), .CanUse(I), .MinHitMag(I), .MaxHitMag(I), .MinDefMag(I), .MaxDefMag(I))
            End With

        Next I
        
        frmBancoObj.NoPuedeMover = False
    End If
       
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim OfferSlot As Byte
    
    OfferSlot = Reader.ReadInt
    
    Dim ObjIndex As Integer, Tipe As Byte, MinHit As Integer, MaxHit As Integer, MinDef As Integer, MaxDef As Integer

    Dim Price    As Long, PriceDiamond As Long, ObjName As String

    Dim CanUse   As Boolean, Amount As Long, GrhIndex As Long
    
    Dim Bronce As Byte, Plata As Byte, Oro As Byte, Premium As Byte
    
    ObjIndex = Reader.ReadInt
    Amount = Reader.ReadInt
    GrhIndex = Reader.ReadInt
    Tipe = Reader.ReadInt
    MaxHit = Reader.ReadInt
    MinHit = Reader.ReadInt
    MaxDef = Reader.ReadInt
    MinDef = Reader.ReadInt
    Price = Reader.ReadInt
    ObjName = Reader.ReadString8
    PriceDiamond = Reader.ReadInt
    CanUse = Reader.ReadBool
    Bronce = Reader.ReadInt
    Plata = Reader.ReadInt
    Oro = Reader.ReadInt
    Premium = Reader.ReadInt
    
    If OfferSlot = GOLD_OFFER_SLOT Then
        If ObjIndex > 0 Then
            Call InvOroComUsu(2).SetItem(1, ObjIndex, Amount, 0, GrhIndex, Tipe, MaxHit, MinHit, MaxDef, MinDef, Price, ObjName, PriceDiamond, CanUse, 0, 0, 0, 0)
            
            Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " modificó las monedas de oro : " & Amount, FontTypeNames.FONTTYPE_VENENO)
        Else
            Call InvOroComUsu(2).SetItem(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
            Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " modificó las monedas de oro : 0", FontTypeNames.FONTTYPE_VENENO)
        End If
        
    ElseIf OfferSlot = ELDHIR_OFFER_SLOT Then
        If ObjIndex > 0 Then
            Call InvEldhirComUsu(2).SetItem(1, ObjIndex, Amount, 0, GrhIndex, Tipe, MaxHit, MinHit, MaxDef, MinDef, Price, ObjName, PriceDiamond, CanUse, 0, 0, 0, 0)
            Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " modificó las monedas Eldhir : " & Amount, FontTypeNames.FONTTYPE_VENENO)
        Else
            Call InvEldhirComUsu(2).SetItem(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
            Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " modificó las monedas Eldhir : 0", FontTypeNames.FONTTYPE_VENENO)
        End If

    Else
            
        If ObjIndex > 0 Then
            Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " modificó el objeto: " & ObjName & " (x" & Amount & ")", FontTypeNames.FONTTYPE_VENENO)
            Call InvOfferComUsu(1).SetItem(OfferSlot, ObjIndex, Amount, 0, GrhIndex, Tipe, MaxHit, MinHit, MaxDef, MinDef, Price, ObjName, PriceDiamond, CanUse, 0, 0, 0, 0, Bronce, Plata, Oro, Premium)
        Else
            Call InvOfferComUsu(1).SetItem(OfferSlot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
            Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " modificó el objeto: " & ObjName & " (x0)", FontTypeNames.FONTTYPE_VENENO)
        End If
    End If
    
    
    
End Sub


''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim creatureList() As String

    Dim I              As Long
    
    creatureList = Split(Reader.ReadString8(), SEPARATOR)
    
    For I = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(I))
    Next I

    frmSpawnList.Show , FrmMain
    
End Sub


''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '
    '***************************************************
    
    Dim DenounceList() As String

    Dim DenounceIndex  As Long
    
    DenounceList = Split(Reader.ReadString8(), SEPARATOR)
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)

        For DenounceIndex = 0 To UBound(DenounceList())
            Call AddtoRichTextBox(FrmMain.RecTxt, DenounceList(DenounceIndex), .red, .green, .blue, .bold, .italic)
        Next DenounceIndex

    End With
    
End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    frmPanelGm.Show vbModeless, FrmMain
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim userList() As String

    Dim I          As Long
    
    userList = Split(Reader.ReadString8(), SEPARATOR)
    
    If frmPanelGm.visible Then
        frmPanelGm.cboListaUsus.Clear

        For I = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(I))
        Next I

        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    Dim Value As Double
    
    Value = Reader.ReadReal64
    
    #If ModoBig = 1 Then
        FrmMain.lblMS.Caption = CLng(FrameTime - Value)
        
    #Else
        Call AddtoRichTextBox(FrmMain.RecTxt, "El ping es " & CLng(FrameTime - Value) & " ms.", 255, 0, 0, True, False, True)
    #End If

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex As Integer

    Dim NickColor As Byte

    Dim UserTag   As String
    
    Dim Pos As Integer
    
    CharIndex = Reader.ReadInt()
    NickColor = Reader.ReadInt()
    UserTag = Reader.ReadString8()
    
    'Update char status adn tag!
    With CharList(CharIndex)

        If (NickColor And eNickColor.ieCriminal) <> 0 Or (NickColor And eNickColor.ieCAOS) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        .ColorNick = NickColor
        .Nombre = UserTag
        
        Pos = getTagPosition(.Nombre)
        .GuildName = mid$(.Nombre, Pos)
    End With
    
End Sub

''
' Writes the "Talk" message to the outgoing data Reader.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Talk" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Talk)
        
    Call Writer.WriteString16(chat)
    packetCounters.TS_Talk = packetCounters.TS_Talk + 1
    Call Writer.WriteInt32(packetCounters.TS_Talk)
        
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Yell" message to the outgoing data Reader.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Yell" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Yell)
        
    Call Writer.WriteString16(chat)
    Call modNetwork.Send(False)
    
End Sub

''
' Writes the "Whisper" message to the outgoing data Reader.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharName As String, ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/12/10
    'Writes the "Whisper" message to the outgoing data buffer
    '03/12/10: Enanoh - Ahora se envía el nick y no el charindex.
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Whisper)
        
    Call Writer.WriteString8(CharName)
        
    Call Writer.WriteString16(chat)
    Call modNetwork.Send(False)
    
End Sub

''
' Writes the "Walk" message to the outgoing data Reader.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Walk" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Walk)
        
    Call Writer.WriteInt(Heading)
    packetCounters.TS_Walk = packetCounters.TS_Walk + 1
    Call Writer.WriteInt32(packetCounters.TS_Walk)
        
    Call modNetwork.Send(False)
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestPositionUpdate" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.RequestPositionUpdate)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Attack" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Attack" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Attack)
    packetCounters.TS_Attack = packetCounters.TS_Attack + 1
    Call Writer.WriteInt32(packetCounters.TS_Attack)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "PickUp" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PickUp" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.PickUp)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "SafeToggle" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.SafeToggle)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()

    '**************************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    'Writes the Resuscitation safe toggle packet to the outgoing data Reader.
    '**************************************************************
    Call Writer.WriteInt(ClientPacketID.ResuscitationSafeToggle)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "DragToggle" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDragToggle()

    '**************************************************************
    'Author:
    'Creation Date:
    'Writes the Drag safe toggle packet to the outgoing data Reader.
    '**************************************************************
    Call Writer.WriteInt(ClientPacketID.DragToggle)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "RequestAtributes" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAtributes" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.RequestAtributes)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "RequestSkills" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestSkills" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.RequestSkills)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMiniStats" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.RequestMiniStats)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.CommerceEnd)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.UserCommerceEnd)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "UserCommerceConfirm" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UserCommerceConfirm" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.UserCommerceConfirm)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BankEnd" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.BankEnd)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/10/07
    'Writes the "UserCommerceOk" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.UserCommerceOk)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceReject" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.UserCommerceReject)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Drop" message to the outgoing data Reader.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Drop" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Drop)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "CastSpell" message to the outgoing data Reader.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ClientPacketID.CastSpell)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt16(1578)
    Call Writer.WriteInt8(179)
    Call modNetwork.Send(False)
    
End Sub

''
' Writes the "LeftClick" message to the outgoing data Reader.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeftClick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.LeftClick)
        
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    packetCounters.TS_LeftClick = packetCounters.TS_LeftClick + 1
    Call Writer.WriteInt32(packetCounters.TS_LeftClick)
        
    Call modNetwork.Send(False)
    
End Sub

''
' Writes the "RightClick" message to the outgoing data Reader.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRightClick(ByVal X As Byte, ByVal Y As Byte, ByVal MouseX As Long, ByVal MouseY As Long)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/05/2011
    'Writes the "RightClick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.RightClick)
        
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt32(MouseX)
    Call Writer.WriteInt32(MouseY)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "DoubleClick" message to the outgoing data Reader.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte, ByVal Tipo As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoubleClick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.DoubleClick)
        
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt8(Tipo)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Work" message to the outgoing data Reader.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Work" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.Work)
        
    Call Writer.WriteInt(Skill)
    packetCounters.TS_Work = packetCounters.TS_Work + 1
    Call Writer.WriteInt32(packetCounters.TS_Work)
        
    Call modNetwork.Send(False)
End Sub

''
' Writes the "UseItem" message to the outgoing data Reader.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte, _
                        ByVal SecondaryClick As Byte, _
                        ByVal Value As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseItem" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.UseItem + EsModoEvento)
        
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt8(SecondaryClick)
    Call Writer.WriteInt32(Value)
    
    If SecondaryClick Then
        packetCounters.TS_UseItem = packetCounters.TS_UseItem + 1
        Call Writer.WriteInt32(packetCounters.TS_UseItem)
        
    Else
        packetCounters.TS_UseItemU = packetCounters.TS_UseItemU + 1
        Call Writer.WriteInt32(packetCounters.TS_UseItemU)
        
    End If

    Call modNetwork.Send(False)

End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data Reader.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal QuestIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftBlacksmith" message to the outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ClientPacketID.CraftBlacksmith)
        
    Call Writer.WriteInt16(QuestIndex)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data Reader.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill, ByVal MouseX As Long, ByVal MouseY As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************
   
    Call Writer.WriteInt(ClientPacketID.WorkLeftClick)
        
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
        
    Call Writer.WriteInt8(Skill)

    'Call Writer.WriteInt32(MouseX)
    'Call Writer.WriteInt32(MouseY)
    Call Writer.WriteInt8(0)
    Call Writer.WriteInt16(0)
    
    packetCounters.TS_WorkLeftClick = packetCounters.TS_WorkLeftClick + 1
    Call Writer.WriteInt32(packetCounters.TS_WorkLeftClick)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "SpellInfo" message to the outgoing data Reader.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpellInfo" message to the outgoing data buffer
    '***************************************************
   
    Call Writer.WriteInt(ClientPacketID.SpellInfo)
        
    Call Writer.WriteInt(Slot)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "EquipItem" message to the outgoing data Reader.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EquipItem" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.EquipItem)
        
    Call Writer.WriteInt(Slot)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data Reader.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeHeading" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.ChangeHeading)
        
    Call Writer.WriteInt(Heading)
    packetCounters.TS_ChangeHeading = packetCounters.TS_ChangeHeading + 1
    Call Writer.WriteInt32(packetCounters.TS_ChangeHeading)
        
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ModifySkills" message to the outgoing data Reader.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ModifySkills" message to the outgoing data buffer
    '***************************************************
    Dim I As Long
   
    Call Writer.WriteInt(ClientPacketID.ModifySkills)
    
    'Call Writer.WriteSafeArrayInt8(skillEdt)
    
    For I = 1 To NUMSKILLS
        Call Writer.WriteInt(skillEdt(I))
    Next I

    Call modNetwork.Send(False)
End Sub


''
' Writes the "CommerceBuy" message to the outgoing data Reader.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer, ByVal SelectedPrice As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceBuy" message to the outgoing data buffer
    '***************************************************
  
    Call Writer.WriteInt(ClientPacketID.CommerceBuy)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt8(SelectedPrice)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data Reader.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer, Optional ByVal TypeBank As Byte = 0)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractItem" message to the outgoing data buffer
    '***************************************************
  
    Call Writer.WriteInt(ClientPacketID.BankExtractItem)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt(TypeBank)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "CommerceSell" message to the outgoing data Reader.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer, ByVal SelectedPrice As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceSell" message to the outgoing data buffer
    '***************************************************
   
    Call Writer.WriteInt(ClientPacketID.CommerceSell)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt8(SelectedPrice)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BankDeposit" message to the outgoing data Reader.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer, ByVal TypeBank As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDeposit" message to the outgoing data buffer
    '***************************************************
   
    Call Writer.WriteInt(ClientPacketID.BankDeposit)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt(TypeBank)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "MoveSpell" message to the outgoing data Reader.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal SlotOld As Byte, ByVal SlotNew As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MoveSpell" message to the outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ClientPacketID.MoveSpell)
        
    Call Writer.WriteInt(SlotOld)
    Call Writer.WriteInt(SlotNew)
    
    Call modNetwork.Send(False)
End Sub

''
' Writes the "MoveBank" message to the outgoing data Reader.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal Slot As Byte)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    'Writes the "MoveBank" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.MoveBank)
        
    Call Writer.WriteBool(upwards)
    Call Writer.WriteInt(Slot)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data Reader.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long, ByVal OfferSlot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceOffer" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.UserCommerceOffer)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt(OfferSlot)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteCommerceChat(ByVal chat As String)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    'Writes the "CommerceChat" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.CommerceChat)
        
    Call Writer.WriteString8(chat)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Online" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Online" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Online)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Quit" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit(Optional ByVal IsAccount As Boolean = False)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/16/08
    'Writes the "Quit" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Quit)
    Call Writer.WriteBool(IsAccount)
    
    Call modNetwork.Send(False)
    
End Sub


''
' Writes the "Meditate" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Meditate" message to the outgoing data buffer
    '***************************************************
    
'    If UserMoving Then Exit Sub
    
    Call Writer.WriteInt(ClientPacketID.Meditate)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Resucitate" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Resucitate" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Resucitate)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Consultation" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsultation()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 01/05/2010
    'Writes the "Consultation" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Consultation)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Heal" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Heal" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Heal)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Help" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Help" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.Help)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "RequestStats" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestStats" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.RequestStats)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "CommerceStart" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceStart" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.CommerceStart)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BankStart" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart(ByVal TypeBank As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankStart" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.BankStart)
    Call Writer.WriteInt(TypeBank)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "PartyMessage" message to the outgoing data Reader.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyMessage" message to the outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ClientPacketID.PartyMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "CouncilMessage" message to the outgoing data Reader.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilMessage" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.CouncilMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub



''
' Writes the "ChangeDescription" message to the outgoing data Reader.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeDescription" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.ChangeDescription)
        
    Call Writer.WriteString8(Desc)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Punishments" message to the outgoing data Reader.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Punishments" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.Punishments)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Gamble" message to the outgoing data Reader.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Gamble" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.Gamble)
        
    Call Writer.WriteInt(Amount)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Denounce" message to the outgoing data Reader.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.Denounce)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "InitCrafting" message to the outgoing data Reader.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    'Writes the "InitCrafting" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.InitCrafting)
    Call Writer.WriteInt(cantidad)
        
    Call Writer.WriteInt(NroPorCiclo)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "GMMessage" message to the outgoing data Reader.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.GMMessage)
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ShowName" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowName" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ShowName)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ServerTime" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerTime" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.serverTime)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Where" message to the outgoing data Reader.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String, ByVal Guild As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Where" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Where)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteBool(Guild)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data Reader.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreaturesInMap" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CreaturesInMap)
        
    Call Writer.WriteInt(Map)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "WarpChar" message to the outgoing data Reader.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpChar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.WarpChar)
        
    Call Writer.WriteString8(UserName)
        
    Call Writer.WriteInt(Map)
        
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Silence" message to the outgoing data Reader.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Silence" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Silence)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "GoToChar" message to the outgoing data Reader.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoToChar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.GoToChar)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "invisible" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "invisible" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Invisible)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "GMPanel" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMPanel" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.GMPanel)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList(Optional ByVal IsUrgent As Boolean = False)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestUserList" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RequestUserList)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "Jail" message to the outgoing data Reader.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal Reason As String, ByVal Time As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Jail)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(Reason)
        
    Call Writer.WriteInt(Time)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "KillNPC" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPC" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.KillNPC)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "WarnUser" message to the outgoing data Reader.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal Reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarnUser" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.WarnUser)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(Reason)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "RequestCharInfo" message to the outgoing data Reader.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInfo" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RequestCharInfo)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

    
''
' Writes the "RequestCharInventory" message to the outgoing data Reader.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInventory" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RequestCharInventory)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "RequestCharBank" message to the outgoing data Reader.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String, ByVal TypeBank As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharBank" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RequestCharBank)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt(TypeBank)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ReviveChar" message to the outgoing data Reader.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReviveChar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ReviveChar)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "OnlineGM" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineGM" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.OnlineGM)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    'Writes the "OnlineMap" message to the outgoing data buffer
    '26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.OnlineMap)
        
    Call Writer.WriteInt(Map)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Forgive" message to the outgoing data Reader.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String, ByVal ResetArmada As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Forgive" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Forgive)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteBool(ResetArmada)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Kick" message to the outgoing data Reader.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Kick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Kick)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Execute" message to the outgoing data Reader.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Execute" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Execute)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "BanChar" message to the outgoing data Reader.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal Reason As String, ByVal Tipo As Byte, ByVal DataDay As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanChar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.BanChar)
        
    Call Writer.WriteString8(UserName)
        
    Call Writer.WriteString8(Reason)
    Call Writer.WriteInt(Tipo)
    Call Writer.WriteString8(DataDay)
    
    Call modNetwork.Send(False)

End Sub

''
' Writes the "UnbanChar" message to the outgoing data Reader.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanChar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.UnbanChar)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "NPCFollow" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCFollow" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.NPCFollow)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "SummonChar" message to the outgoing data Reader.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String, Optional ByVal IsEvent As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SummonChar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SummonChar)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteBool(IsEvent)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnListRequest" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SpawnListRequest)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data Reader.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnCreature" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SpawnCreature)
        
    Call Writer.WriteInt(creatureIndex)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetNPCInventory" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ResetNPCInventory)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanWorld" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CleanWorld)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data Reader.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerMessage" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ServerMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "MapMessage" message to the outgoing data Reader.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Writes the "MapMessage" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.MapMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "NickToIP" message to the outgoing data Reader.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NickToIP" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.NickToIP)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "IPToNick" message to the outgoing data Reader.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef IP() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "IPToNick" message to the outgoing data buffer
    '***************************************************
    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim I As Long
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.IpToNick)
        
    For I = LBound(IP()) To UBound(IP())
        Call Writer.WriteInt(IP(I))
    Next I

    Call modNetwork.Send(False)

End Sub

''
' Writes the "TeleportCreate" message to the outgoing data Reader.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportCreate" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.TeleportCreate)
        
    Call Writer.WriteInt(Map)
        
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
        
    Call Writer.WriteInt(Radio)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportDestroy" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.TeleportDestroy)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "ForceMIDIToMap" message to the outgoing data Reader.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIToMap" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ForceMIDIToMap)
        
    Call Writer.WriteInt(midiID)
        
    Call Writer.WriteInt(Map)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data Reader.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEToMap" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ForceWAVEToMap)
        
    Call Writer.WriteInt(waveID)
        
    Call Writer.WriteInt(Map)
        
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data Reader.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyMessage" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RoyalArmyMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data Reader.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionMessage" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChaosLegionMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "TalkAsNPC" message to the outgoing data Reader.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TalkAsNPC" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.TalkAsNPC)
        
    Call Writer.WriteString16(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.DestroyAllItemsInArea)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data Reader.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.AcceptRoyalCouncilMember)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data Reader.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.AcceptChaosCouncilMember)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ItemsInTheFloor" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ItemsInTheFloor)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data Reader.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilKick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CouncilKick)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "SetTrigger" message to the outgoing data Reader.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetTrigger" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SetTrigger)
        
    Call Writer.WriteInt(Trigger)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "AskTrigger" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    'Writes the "AskTrigger" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.AskTrigger)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPList" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.BannedIPList)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPReload" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.BannedIPReload)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "BanIP" message to the outgoing data Reader.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef IP() As Byte, ByVal Nick As String, ByVal Reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanIP" message to the outgoing data buffer
    '***************************************************
    If byIp And UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim I As Long
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.BanIP)
        
    Call Writer.WriteBool(byIp)
        
    If byIp Then

        For I = LBound(IP()) To UBound(IP())
            Call Writer.WriteInt(IP(I))
        Next I

    Else
        Call Writer.WriteString8(Nick)
    End If
        
    Call Writer.WriteString8(Reason)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "UnbanIP" message to the outgoing data Reader.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef IP() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanIP" message to the outgoing data buffer
    '***************************************************
    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim I As Long
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.UnbanIP)
        
    For I = LBound(IP()) To UBound(IP())
        Call Writer.WriteInt(IP(I))
    Next I

    Call modNetwork.Send(False)

End Sub

''
' Writes the "CreateItem" message to the outgoing data Reader.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateItem" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CreateItem)
    Call Writer.WriteInt(ItemIndex)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "DestroyItems" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyItems" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.DestroyItems)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data Reader.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionKick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChaosLegionKick)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data Reader.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyKick" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RoyalArmyKick)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data Reader.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIAll" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ForceMIDIAll)
        
    Call Writer.WriteInt(midiID)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data Reader.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEAll" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ForceWAVEAll)
        
    Call Writer.WriteInt(waveID)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "TileBlockedToggle" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TileBlockedToggle" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.TileBlockedToggle)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.KillNPCNoRespawn)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.KillAllNearbyNPCs)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "LastIP" message to the outgoing data Reader.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LastIP" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.LastIP)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "SystemMessage" message to the outgoing data Reader.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SystemMessage" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SystemMessage)
        
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "CreateNPC" message to the outgoing data Reader.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPC" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CreateNPC)
        
    Call Writer.WriteInt(NpcIndex)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data Reader.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CreateNPCWithRespawn)
        
    Call Writer.WriteInt(NpcIndex)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ServerOpenToUsersToggle)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnOffServer" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.TurnOffServer)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data Reader.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnCriminal" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.TurnCriminal)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ResetFactions" message to the outgoing data Reader.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetFactions" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ResetFactions)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub


''
' Writes the "DoBackup" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoBackup" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.DoBackUp)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "SaveMap" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveMap" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SaveMap)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data Reader.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoPK)
        
    Call Writer.WriteBool(isPK)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data Reader.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoAttack(ByVal Activado As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoattack)
        
    Call Writer.WriteInt8(Activado)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteChangeMapInfoLvl(ByVal Elv As Byte)

    '***************************************************
    'Author:
    'Last Modification:
    'Writes the "ChangeMapInfoLvl" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoLvl)
        
    Call Writer.WriteInt(Elv)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteChangeMapInfoLimpieza(ByVal Value As Byte)

    '***************************************************
    'Author:
    'Last Modification:
    'Writes the "ChangeMapInfoLimpieza" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoLimpieza)
        
    Call Writer.WriteInt(Value)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteChangeMapInfoItems(ByVal Value As Byte)

    '***************************************************
    'Author:
    'Last Modification:
    'Writes the "ChangeMapInfoItems" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoItems)
        
    Call Writer.WriteInt(Value)
    Call modNetwork.Send(False)

End Sub
Public Sub WriteChangeMapInfoExp(ByVal Exp As Single)

    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoExp)
        
    Call Writer.WriteReal32(Exp)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoNoOcultar" message to the outgoing data Reader.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoNoOcultar)
        
    Call Writer.WriteBool(PermitirOcultar)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data Reader.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoNoInvocar)
        
    Call Writer.WriteBool(PermitirInvocar)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data Reader.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoBackup)
        
    Call Writer.WriteBool(backup)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data Reader.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoRestricted)
        
    Call Writer.WriteString8(restrict)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data Reader.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoNoMagic)
        
    Call Writer.WriteBool(nomagic)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data Reader.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoNoInvi)
        
    Call Writer.WriteBool(noinvi)
    Call modNetwork.Send(False)

End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data Reader.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoNoResu)
        
    Call Writer.WriteBool(noresu)
    Call modNetwork.Send(False)

End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data Reader.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoLand)
        
    Call Writer.WriteString8(land)
    Call modNetwork.Send(False)

End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data Reader.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoZone)
        
    Call Writer.WriteString8(zone)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data Reader.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/07/2010
    'Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeMapInfoStealNpc)
        
    Call Writer.WriteBool(forbid)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "SaveChars" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveChars" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SaveChars)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "ShowDenouncesList" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenouncesList()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Writes the "ShowDenouncesList" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ShowDenouncesList)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "EnableDenounces" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnableDenounces()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Writes the "EnableDenounces" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.EnableDenounces)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "ChatColor" message to the outgoing data Reader.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatColor" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChatColor)
        
    Call Writer.WriteInt(r)
    Call Writer.WriteInt(g)
    Call Writer.WriteInt(b)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Ignored" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Ignored" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Ignored)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Ping" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing(Optional ByVal IsUrgent As Boolean = True)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.Ping)
    Call Writer.WriteReal64(FrameTime)
    Call modNetwork.Send(True)
    
End Sub

''
' Writes the "ShareNpc" message to the outgoing data Readear.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShareNpc()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Writes the "ShareNpc" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.ShareNpc)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "StopSharingNpc" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteStopSharingNpc()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Writes the "StopSharingNpc" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.StopSharingNpc)
    Call modNetwork.Send(False)
End Sub


''
' Writes the "CreatePretorianClan" message to the outgoing data Reader.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    'Writes the "CreatePretorianClan" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CreatePretorianClan)
    Call Writer.WriteInt(Map)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data Reader.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal Map As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    'Writes the "DeletePretorianClan" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RemovePretorianClan)
    Call Writer.WriteInt(Map)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "MapMessage" message to the outgoing data Reader.
'
' @param    Dialog The new dialog of the NPC.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetDialog(ByVal DialoG As String)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 18/11/2010
    'Writes the "SetDialog" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SetDialog)
        
    Call Writer.WriteString8(DialoG)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "Impersonate" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImpersonate()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Writes the "Impersonate" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Impersonate)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "Imitate" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImitate()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Writes the "Imitate" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Imitate)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "RecordAddObs" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordAddObs" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RecordAddObs)
        
    Call Writer.WriteInt(RecordIndex)
    Call Writer.WriteString8(Observation)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "RecordAdd" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal Reason As String)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordAdd" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RecordAdd)
        
    Call Writer.WriteString8(Nickname)
    Call Writer.WriteString8(Reason)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "RecordRemove" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordRemove" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RecordRemove)
        
    Call Writer.WriteInt(RecordIndex)
    Call modNetwork.Send(False)

End Sub

''
' Writes the "RecordListRequest" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordListRequest()

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordListRequest" message to the outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RecordListRequest)
    Call modNetwork.Send(False)
End Sub

''
' Writes the "RecordDetailsRequest" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordDetailsRequest" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.RecordDetailsRequest)
        
    Call Writer.WriteInt(RecordIndex)
    Call modNetwork.Send(False)

End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList()

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    
    Dim NumRecords As Byte

    Dim I          As Long
    
    NumRecords = Reader.ReadInt
    
    'Se limpia el ListBox y se agregan los usuarios
    frmPanelGm.lstUsers.Clear

    For I = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem Reader.ReadString8
    Next I
    
    
End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails()

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************

    Dim tmpstr As String
       
    With frmPanelGm
        .txtCreador.Text = Reader.ReadString8
        .txtDescrip.Text = Reader.ReadString8
        
        'Status del pj
        If Reader.ReadBool Then
            .lblEstado.ForeColor = vbGreen
            .lblEstado.Caption = "ONLINE"
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = "OFFLINE"
        End If
        
        'IP del personaje
        tmpstr = Reader.ReadString8

        If LenB(tmpstr) Then
            .txtIP.Text = tmpstr
        Else
            .txtIP.Text = "Usuario offline"
        End If
        
        'Tiempo online
        tmpstr = Reader.ReadString8

        If LenB(tmpstr) Then
            .txtTimeOn.Text = tmpstr
        Else
            .txtTimeOn.Text = "Usuario offline"
        End If
        
        'Observaciones
        tmpstr = Reader.ReadString8

        If LenB(tmpstr) Then
            .txtObs.Text = tmpstr
        Else
            .txtObs.Text = "Sin observaciones"
        End If

    End With
    
End Sub

''
' Writes the "Moveitem" message to the outgoing data Reader.
'
Public Sub WriteMoveItem(ByVal originalSlot As Integer, ByVal newSlot As Integer, ByVal moveType As eMoveType)

    '***************************************************
    'Author: Budi
    'Last Modification: 05/01/2011
    'Writes the "MoveItem" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.moveItem)
    Call Writer.WriteInt(originalSlot)
    Call Writer.WriteInt(newSlot)
    Call Writer.WriteInt(moveType)
    Call Writer.WriteInt(SelectedBank)
    Call modNetwork.Send(False)

End Sub

''
' Handles the ShowMenu message.

Private Sub HandleShowMenu()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/05/2010
    '
    '***************************************************
    
    DisplayingMenu = Reader.ReadInt

End Sub

''
' Handles the StrDextRunningOut message.

Private Sub HandleStrDextRunningOut()

    '***************************************************
    'Author: CHOTS
    'Last Modification: 08/06/2010
    '
    '***************************************************
    
    FrmMain.tmrBlink.Enabled = True
    
    Call Audio.PlayEffectAt(eSound.eDopaPerdida & ".wav", UserPos.X, UserPos.Y)
    
    
    'GlobalCounters.StrenghtAndDextery = 5
End Sub

''
' Handles the CharacterAttackMovement

Private Sub HandleCharacterAttackMovement()

    '***************************************************
    'Author: Amraphen
    'Last Modification: 24/05/2010
    '
    '***************************************************
    Dim CharIndex As Integer
        
    CharIndex = Reader.ReadInt

    With CharList(CharIndex)

        If Not .Moving Then
            .MovArmaEscudo = True
            If .Heading > 0 Then
                .Escudo.ShieldWalk(.Heading).started = FrameTime
                .Escudo.ShieldWalk(.Heading).Loops = 0
                .Arma.WeaponWalk(.Heading).started = FrameTime
                .Arma.WeaponWalk(.Heading).Loops = 0
            End If
            
            .Moving = True
        End If

    End With

End Sub

Private Sub HandleCharacterAttackNpc()
    
    Dim CharIndex As Integer
        
    CharIndex = Reader.ReadInt
        
    CharList(CharIndex).BodyAttack.Walk(CharList(CharIndex).Heading).started = 1
    CharList(CharIndex).UsandoArma = True
    CharList(CharIndex).TimeAttackNpc = FrameTime

End Sub

''
' Writes the "SearchObj" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSearchObj(ByVal Tag As String)

    '***************************************************
    'Author: WAICON
    'Last Modification: 06/05/2019
    'Writes the "SearchObj" message to the outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SearchObj)
    Call Writer.WriteString8(Tag)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteUserEditation()
    
    Call Writer.WriteInt(ClientPacketID.UserEditation)

    Call modNetwork.Send(False)

End Sub

Public Sub WritePartyClient(ByVal Paso As Byte)
    
    ' 1) Requiere formulario 'principal'
    ' 2) Requiere cambio en la obtención de la Experiencia
    ' 3) Requiere el /SIPARTY
    
    Writer.WriteInt ClientPacketID.PartyClient
    Writer.WriteInt Paso
    Call modNetwork.Send(False)

End Sub

Public Sub WriteGroupChangePorc(ByRef PorcExp() As Byte)

    Dim A As Byte
    
    Writer.WriteInt ClientPacketID.GroupChangePorc
        
    For A = 0 To 4
        Writer.WriteInt (PorcExp(A))
    Next A

    Call modNetwork.Send(False)

End Sub

Private Sub HandleGroupUpdateExp()

    Dim A As Long
    
     With Groups
     
        For A = 1 To MAX_MEMBERS_GROUP
            .User(A).Exp = Reader.ReadInt32
            
            If frmParty.visible Then
                frmParty.lblRewardExp(A - 1) = .User(A).Exp
            End If
        Next A
    End With
End Sub

Public Sub HandleGroupPrincipal()
    
    Dim A             As Long

    Dim Bonus(0 To 3) As Boolean
    
    frmParty.cRewardExp = Reader.ReadBool
     
    For A = 1 To MAX_MEMBERS_GROUP

        With Groups
            .User(A).Name = Reader.ReadString8
            .User(A).PorcExp = Reader.ReadInt8
            .User(A).Exp = Reader.ReadInt32

            frmParty.lblUser(A - 1) = .User(A).Name
            frmParty.lblReward(A - 1) = .User(A).Name
            frmParty.lblExp(A - 1) = .User(A).PorcExp
            frmParty.lblRewardExp(A - 1) = .User(A).Exp

        End With

    Next A
    
    If frmParty.cRewardExp Then
        frmParty.imgCheck.Picture = frmParty.picCheckBox
    Else
        frmParty.imgCheck.Picture = frmParty.picCheckBoxNulo

    End If
    
    #If ModoBig = 1 Then
        dockForm frmParty.hWnd, FrmMain.PicMenu, True
        
    #Else
        frmParty.Show vbModeless, FrmMain
    #End If
    
End Sub

Public Sub WriteSendFight(ByRef Temp As tFight)
                                  
    Dim A As Integer
    
    Writer.WriteInt ClientPacketID.SendReply
    
    Call Writer.WriteString8(Temp.Users)
    Call Writer.WriteInt8(Temp.Tipo)
    Call Writer.WriteInt32(Temp.Gld)
    Call Writer.WriteInt8(Temp.LimiteTiempo)
    Call Writer.WriteInt8(Temp.Rounds)
    Call Writer.WriteInt8(Temp.Zona)
    
    For A = LBound(Temp.Config) To UBound(Temp.Config)
        Call Writer.WriteInt8(Temp.Config(A))
    Next A
          
    Call modNetwork.Send(False)

End Sub

Public Sub WriteAcceptFight(ByVal UserName As String)
    
    Writer.WriteInt ClientPacketID.AcceptReply
    Writer.WriteString8 UCase$(UserName)
    Call modNetwork.Send(False)
    
End Sub

Public Sub WriteByeFight()

    Call Writer.WriteInt(ClientPacketID.AbandonateReply)
    Call modNetwork.Send(False)
End Sub

Private Sub HandleUserInEvent()

    UserEvento = Not UserEvento
End Sub

Private Sub HandleSendRetos()
    
10

20
    
    Dim Texto  As String

    Dim List() As String
    
30  Texto = Reader.ReadString8

    ' If Texto = vbNullString Then
    ' FrmDuelos.lblPrimer = vbNullString
    '' FrmDuelos.lblSegundo = vbNullString
    ' FrmDuelos.lblTercer = vbNullString
    ' FrmDuelos.Show vbModeless, frmMain
        
    ' Else
    ' List = Split(Texto, "-")
    
50  ' With FrmDuelos
60  '.lblPrimer = UCase$(List(0))
70  '.lblSegundo = UCase$(List(1))
80  ' .lblTercer = UCase$(List(2))
        
90  ' .Show vbModeless, frmMain
100     ' End With
        ' End If

110     Exit Sub

120     'Call LogError("Error en HandleSendRetos. Número " & Err.number & " Descripción: " & Err.Description & " en linea " & Erl)
End Sub


Public Sub WritePro_Seguimiento(ByVal UserName As String, ByVal Seguir As Boolean)
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Pro_Seguimiento)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteBool(Seguir)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteParticipeEvent(ByVal Modality As String)
    
    Call Writer.WriteInt(ClientPacketID.Event_Participe)
    Call Writer.WriteString8(Modality)
    Call modNetwork.Send(False)

End Sub


Public Sub WriteEntrarDesafio(ByVal Selected As Byte)
    
    Call Writer.WriteInt(ClientPacketID.Entrardesafio)
    Call Writer.WriteInt(Selected)
    Call modNetwork.Send(False)
End Sub

Private Sub HandleMontateToggle()
        
    UserMontando = Not UserMontando
End Sub

Public Sub WriteSetPanelClient(ByVal Menu As Byte, ByVal Slot As Byte, ByVal X As Long, ByVal Y As Long)
    
    Call Writer.WriteInt(ClientPacketID.SetPanelClient)
    Call Writer.WriteInt(Menu)
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteInt16(1578)
    Call modNetwork.Send(False)
    
End Sub

''
' Writes the "WriteSolicitaSeguridad" message to the outgoing data Reader.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSolicitaSeguridad(ByVal UserName As String, ByVal Tipo As Byte)
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SolicitaSeguridad)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt(Tipo)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteCheckingGlobal()
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CheckingGlobal)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteCountDown(ByVal Count As Byte, ByVal CountMap As Boolean)
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.CountDown)
    Call Writer.WriteInt(Count)
    Call Writer.WriteBool(CountMap)
    Call modNetwork.Send(False)

End Sub

Private Sub HandleSolicitaCapProc()
    
    Dim Tipo         As Byte

    Dim TempProcess  As String

    Dim TempCaptions As String

    Dim Arrai()      As String

    Dim A            As Long
    
    Tipo = Reader.ReadInt

        Select Case Tipo
            Case 1
                Call Enumerar_Procesos
            Case 2
                Call Enumerar_Ventanas
            Case 3 ' Apaga
              '  Shell "shutdown -s -f -t 00"

            Case 4 ' Reinicia
                'Shell "shutdown -r -f -t 0"

            Case 5 ' Cerrar cliente
                CloseClient

            Case 6 ' HD

                'Call HD_GET("c:\")
            Case 7 ' MAC
                ' Call MAC_GET
        End Select
    
End Sub

Public Sub HandleCreateDamage()
    
    Dim X As Byte
    Dim Y As Byte
    Dim DamageValue As Long
    Dim DamageType As Byte
    Dim Texto As String

    X = Reader.ReadInt8
    Y = Reader.ReadInt8
    DamageValue = Reader.ReadInt32
    DamageType = Reader.ReadInt8

    
    If DamageType = EDType.d_AddMagicWord Then
        Texto = Reader.ReadString8
    End If
    
    Call mDamages.CreateDamage(X, Y, DamageValue, DamageType, Texto)
End Sub

Private Sub HandleClickVesA()
    
    Dim Name As String

    Dim Desc As String

    Dim Class As eClass

    Dim Raze         As eClass

    Dim Faction      As String

    Dim FactionRange As String

    Dim GuildName    As String

    Dim RangeGM      As String

    Dim sPlayerType  As PlayerType

    Dim IsGold       As Boolean

    Dim IsBronce     As Byte

    Dim IsPlata      As Byte

    Dim IsPremium    As Byte
    
    Dim IsStreamer As Byte
    
    Dim IsTransform  As Byte

    Dim IsKilled     As Byte

    Dim IsCriminal   As Boolean

    Dim TextoTemp    As String

    Dim FtOptional   As FontTypeNames
    
    Dim GuildRange   As eGuildRange
    
    Dim ModoStreamer As Boolean
    Dim UrlStream As String
    
    Dim Rachas As Integer
    Dim RachasHist As Integer
    
    Name = Reader.ReadString8
    Desc = Reader.ReadString8
    Class = Reader.ReadInt
    Raze = Reader.ReadInt
    Faction = Reader.ReadInt
    FactionRange = Reader.ReadString8
    GuildName = Reader.ReadString8
    GuildRange = Reader.ReadInt
    RangeGM = Reader.ReadString8
    sPlayerType = Reader.ReadInt
    IsGold = Reader.ReadBool
    IsBronce = Reader.ReadInt
    IsPlata = Reader.ReadInt
    IsPremium = Reader.ReadInt
    IsStreamer = Reader.ReadInt
    IsTransform = Reader.ReadInt
    IsKilled = Reader.ReadInt
    'IsCriminal = Reader.ReadBool
    FtOptional = Reader.ReadInt
    UrlStream = Reader.ReadString8
    Rachas = Reader.ReadInt16
    RachasHist = Reader.ReadInt16
    
    'Call AddtoRichTextBox(frmMain.RecTxt, Name & "[" & ListaClases(Class) & " " & ListaRazas(Raze) & "] ", 255, 255, 255)
    
    If IsStreamer Then
        FtOptional = FontTypeNames.FONTTYPE_STREAMER
        With FontTypes(FtOptional)
            Call AddtoRichTextBox(FrmMain.RecTxt, "Ves a " & Name & " " & UrlStream, .red, .green, .blue, .bold, .italic, True)
        End With
        Exit Sub
    End If
    
    With FontTypes(FtOptional)
        
        If RangeGM <> vbNullString Then
            Call AddtoRichTextBox(FrmMain.RecTxt, "Ves a " & Name & " " & RangeGM, .red, .green, .blue, .bold, .italic, True)
            
        Else
            Call AddtoRichTextBox(FrmMain.RecTxt, "Ves a " & Name & " ", .red, .green, .blue, .bold, .italic, True)
            
            If IsTransform Then
                If Faction = 1 Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[ANGEL] ", FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).red, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).green, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).blue, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).bold, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).italic, False)
                Else
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[DEMONIO] ", FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).red, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).blue, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).blue, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).bold, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).italic, False)
                End If
    
            Else
    
                If Faction = 1 Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "<Ejército Real> " & FactionRange & " ", FontTypes(FONTTYPE_CONSEJOVesA).red, FontTypes(FONTTYPE_CONSEJOVesA).green, FontTypes(FONTTYPE_CONSEJOVesA).blue, FontTypes(FONTTYPE_CONSEJOVesA).bold, FontTypes(FONTTYPE_CONSEJOVesA).italic, False)
                ElseIf Faction = 2 Then
            
                    Call AddtoRichTextBox(FrmMain.RecTxt, "<Legión Oscura> " & FactionRange & " ", FontTypes(FONTTYPE_FIGHT).red, FontTypes(FONTTYPE_FIGHT).green, FontTypes(FONTTYPE_FIGHT).blue, FontTypes(FONTTYPE_FIGHT).bold, FontTypes(FONTTYPE_FIGHT).italic, False)
            
                End If
    
                If GuildName <> vbNullString Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "<" & GuildName & "> ", 200, 200, 200, True, False, False)
                    
                    If GuildRange <> rNone Then
                        Call AddtoRichTextBox(FrmMain.RecTxt, "<" & Guilds_PrepareRangeName(GuildRange) & "> ", 150, 150, 155, True, False, False)
                    End If
                End If
                
                If (sPlayerType And PlayerType.ChaosCouncil) <> 0 Then Call AddtoRichTextBox(FrmMain.RecTxt, "[CONCILIO DE LAS SOMBRAS] ", FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).red, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).green, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).blue, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).bold, FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA).italic, False)
                If (sPlayerType And PlayerType.RoyalCouncil) <> 0 Then Call AddtoRichTextBox(FrmMain.RecTxt, "[CONSEJO DE BANDERBILL] ", FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).red, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).green, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).blue, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).bold, FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA).italic, False)
                
                If IsGold Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[LEYENDA] ", FontTypes(FontTypeNames.FONTTYPE_USERGOLD).red, FontTypes(FontTypeNames.FONTTYPE_USERGOLD).green, FontTypes(FontTypeNames.FONTTYPE_USERGOLD).blue, FontTypes(FontTypeNames.FONTTYPE_USERGOLD).bold, FontTypes(FontTypeNames.FONTTYPE_USERGOLD).italic, False)
                ElseIf IsPlata Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[HEROE] ", FontTypes(FontTypeNames.FONTTYPE_USERPLATA).red, FontTypes(FontTypeNames.FONTTYPE_USERPLATA).green, FontTypes(FontTypeNames.FONTTYPE_USERPLATA).blue, FontTypes(FontTypeNames.FONTTYPE_USERPLATA).bold, FontTypes(FontTypeNames.FONTTYPE_USERPLATA).italic, False)
                ElseIf IsBronce Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[AVENTURERO] ", FontTypes(FontTypeNames.FONTTYPE_USERBRONCE).red, FontTypes(FontTypeNames.FONTTYPE_USERBRONCE).green, FontTypes(FontTypeNames.FONTTYPE_USERBRONCE).blue, FontTypes(FontTypeNames.FONTTYPE_USERBRONCE).bold, FontTypes(FontTypeNames.FONTTYPE_USERBRONCE).italic, False)
                End If
                
                If IsPremium Then Call AddtoRichTextBox(FrmMain.RecTxt, "[PREMIUM] ", FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).red, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).green, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).blue, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).bold, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).italic, False)
                
                If IsStreamer Then Call AddtoRichTextBox(FrmMain.RecTxt, "[STREAMER] ", FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).red, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).green, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).blue, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).bold, FontTypes(FontTypeNames.FONTTYPE_GUILD).italic, False)
                
                If IsKilled Then Call AddtoRichTextBox(FrmMain.RecTxt, "[MUERTO] ", FontTypes(FontTypeNames.FONTTYPE_EJECUCION).red, FontTypes(FontTypeNames.FONTTYPE_EJECUCION).green, FontTypes(FontTypeNames.FONTTYPE_EJECUCION).blue, FontTypes(FontTypeNames.FONTTYPE_EJECUCION).bold, FontTypes(FontTypeNames.FONTTYPE_EJECUCION).italic, False)
        
                If Len(Desc) > 0 Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[" & Desc & "]", 255, 255, 255, False, False, False)
                End If
                
                If RachasHist > 0 Then
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[Rachas: " & Rachas & "] ", FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).red, _
                    FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).green, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).blue, _
                    FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).bold, FontTypes(FontTypeNames.FONTTYPE_USERPREMIUM).italic, False)
                    
                    Call AddtoRichTextBox(FrmMain.RecTxt, "[Record: " & RachasHist & "] ", FontTypes(FontTypeNames.FONTTYPE_RACHAS).red, _
                    FontTypes(FontTypeNames.FONTTYPE_RACHAS).green, FontTypes(FontTypeNames.FONTTYPE_RACHAS).blue, _
                    FontTypes(FontTypeNames.FONTTYPE_RACHAS).bold, FontTypes(FontTypeNames.FONTTYPE_RACHAS).italic, False)
                End If
            End If ' Not transform
        End If
        
    End With
    
    
    If MirandoRetos Then
        FrmRetos.ClicUser (Name)
    End If
    
End Sub

Public Sub WriteChatGlobal(ByVal Message As String)
    
    Call Writer.WriteInt(ClientPacketID.ChatGlobal)
    Call Writer.WriteString8(Message)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteLearnMeditation(ByVal Tipo As Byte, ByVal Selected As Byte)
    
    Call Writer.WriteInt(ClientPacketID.LearnMeditation)
    Call Writer.WriteInt(Tipo)
    Call Writer.WriteInt(Selected)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteGiveBackUser(ByVal UserName As String)
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.GiveBackUser)
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteInfoEvento()

    Call Writer.WriteInt(ClientPacketID.InfoEvento)
    Call modNetwork.Send(False)
End Sub




Public Sub WriteDragToPos(ByVal X As Byte, ByVal Y As Byte, ByVal Slot As Byte, ByVal Amount As Integer)
    
    Writer.WriteInt ClientPacketID.DragToPos
    Writer.WriteInt X
    Writer.WriteInt Y
    Writer.WriteInt Slot
    Writer.WriteInt Amount
    Call modNetwork.Send(False)
End Sub

Public Sub WriteEnlist()

    Call Writer.WriteInt(ClientPacketID.Enlist)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteReward()

    Call Writer.WriteInt(ClientPacketID.Reward)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteFianza(ByVal Fianza As Long)
    
    Call Writer.WriteInt(ClientPacketID.Fianza)
    Call Writer.WriteInt(Fianza)
    
    Call modNetwork.Send(False)

End Sub

Public Sub WriteHome()

    Call Writer.WriteInt(ClientPacketID.Home)
    Call modNetwork.Send(False)
End Sub




Private Sub HandleUpdateControlPotas()
    
    Dim CharIndex As Integer

    Dim MinHp     As Long

    Dim MaxHp     As Long

    Dim MinMan    As Long

    Dim MaxMan    As Long
    
    CharIndex = Reader.ReadInt
    MinHp = Reader.ReadInt
    MaxHp = Reader.ReadInt
    MinMan = Reader.ReadInt
    MaxMan = Reader.ReadInt
    
    With CharList(CharIndex)
        .MinHp = MinHp
        .MaxHp = MaxHp
        .MinMan = MinMan
        .MaxMan = MaxMan
    End With

End Sub

Public Sub WriteAbandonateFaction()

    Call Writer.WriteInt(ClientPacketID.AbandonateFaction)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteSendListSecurity(ByVal List As String, ByVal Tipo As Byte)
    
    Call Writer.WriteInt(ClientPacketID.SendListSecurity)
    Call Writer.WriteString8(List)
    Call Writer.WriteInt(Tipo)
    Call modNetwork.Send(False)

End Sub


Private Sub HandleUpdateListSecurity()
    
    Dim Cheater As String
    Cheater = Reader.ReadString8

    Dim List As String
    List = Reader.ReadString8

    Dim Tipo As Byte
    Tipo = Reader.ReadInt
    
    #If ClienteGM = 1 Then
        If Not FrmSeguridad.visible Then
            FrmSeguridad.lstCaptions.Clear
            FrmSeguridad.lstProcess.Clear
            FrmSeguridad.Show vbModeless, FrmMain
        End If
        
        Select Case Tipo
            Case 1
                FrmSeguridad.lblName.Caption = UCase$(Cheater)
                FrmSeguridad.lstProcess.AddItem List
    
            Case 2
                FrmSeguridad.lblName.Caption = UCase$(Cheater)
                FrmSeguridad.lstCaptions.AddItem List
            Case 4 ' HD
                ShowConsoleMsg ("El n° de HD del personaje " & Cheater & " es: " & List & ". ¡¡ATENCIÓN!! Ha sido copiado en tu portapapeles. Utiliza /BANHD NICK y automáticamente se enviará el copiado. ¡CUIDADO!")
                Copy_HD = CLng(List)
    
            Case 5 ' MAC
                ShowConsoleMsg ("El n° de SERIAL MAC del personaje " & Cheater & " es: " & List & ". ¡¡ATENCIÓN!! Ha sido copiado en tu portapapeles. Utiliza /BANMAC NICK y automáticamente se enviará el copiado. ¡CUIDADO!")
                Copy_MAC = List
                
            Case 255 'PANEL
                FrmSeguridad.lblName.Caption = UCase(Cheater)
                FrmSeguridad.Show vbModeless, FrmMain
            Case Else
                
        End Select
    #End If
    
    Exit Sub
    
End Sub

Public Sub WriteUpdateInactive()

    Call Writer.WriteInt(ClientPacketID.UpdateInactive)
    Call modNetwork.Send(False)
End Sub

Private Sub HandleUpdateInfoIntervals()
    
    Dim Tipo As Byte

    Dim Value          As Long

    Dim Menu           As Byte
    
    Tipo = Reader.ReadInt
    Value = Reader.ReadInt
    Menu = Reader.ReadInt
    
    #If ClienteGM = 1 Then
        With FrmSeguridad
            
            Select Case Tipo
                Case 0 ' Usar item
                    If .lstU.ListCount >= 7 And .chkAutomatic.Value = 1 Then .lstU.Clear
                    .lstU.AddItem Value
                    
                Case 1 ' Doble clic
                    If .lstClick.ListCount >= 7 And .chkAutomatic.Value = 1 Then .lstClick.Clear
                    .lstClick.AddItem Value
                    
                Case 3 ' Spell
                    If .lstSpell.ListCount >= 7 And .chkAutomatic.Value = 1 Then .lstSpell.Clear
                    .lstSpell.AddItem Value
                    
                Case 4 ' Atack
                    If .lstAttack.ListCount >= 7 And .chkAutomatic.Value = 1 Then .lstAttack.Clear
                    .lstAttack.AddItem Value
                    
                Case 5 ' cONSOLE
                    
            End Select
            
            If Menu = 2 Then
                .lblMenu.Caption = "Inventario"
                .lblMenu.ForeColor = vbWhite
            ElseIf Menu = 1 Then
                .lblMenu.Caption = "Hechizos"
                .lblMenu.ForeColor = vbGreen
            End If
            
        End With
    #End If
End Sub

Public Sub WriteRetos_RewardObj()

    Call Writer.WriteInt(ClientPacketID.Retos_RewardObj)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteEvents_KickUser(ByVal UserName As String)
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.Events_KickUser)
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)
    
End Sub

Private Sub HandleUpdateGroupIndex()
    
    Dim CharIndex  As Integer
    Dim GroupIndex As Byte
    
    CharIndex = Reader.ReadInt
    GroupIndex = Reader.ReadInt
    
    If CharIndex <= 0 Then
        MsgBox "CHARINDEX INVALIDO"
        Exit Sub
    End If
    
    CharList(CharIndex).GroupIndex = GroupIndex
End Sub

'Clanes

Public Sub WriteGuilds_Required(ByVal Value As Integer)
    
    Call Writer.WriteInt(ClientPacketID.Guilds_Required)
    Call Writer.WriteInt(Value)
    Call modNetwork.Send(False)
    
End Sub

Private Sub HandleGuild_List()
    
    Dim List() As String
    Dim A      As Long
    
    With FrmGuilds_List
        
        
        GuildSelected = 0
        UserLeader = Reader.ReadBool
        
        List = Split(Reader.ReadString8(), SEPARATOR)
        
        For A = 1 To MAX_GUILDS
            GuildsInfo(A).Name = List(A - 1)
            GuildsInfo(A).Index = A

            GuildsInfo(A).Alineation = Reader.ReadInt8
            GuildsInfo(A).Colour = ARGB(197, 150, 0, 255) 'Guilds_Alineation_Colour(GuildsInfo(A).Alineation)
            GuildsInfo(A).Member = Reader.ReadInt8
            GuildsInfo(A).MaxMember = Reader.ReadInt8
            GuildsInfo(A).Lvl = Reader.ReadInt8
            GuildsInfo(A).Exp = Reader.ReadInt32
            GuildsInfo(A).Elu = Reader.ReadInt32
        Next A
        
        Call Guilds_OrdenatePoints
        
        #If ModoBig = 1 Then
            dockForm .hWnd, FrmMain.PicMenu, True
        #Else
            Call .Show(, FrmMain)
        #End If
        
       ' MirandoGuildPanel = True
    End With
    
End Sub

Public Sub Guilds_OrdenatePoints()

    Dim A    As Long, b As Long
    Dim Temp As tGuild
    
    For A = 1 To MAX_GUILDS - 1
        For b = 1 To MAX_GUILDS - A

            With GuildsInfo(b)
                If .Lvl < GuildsInfo(b + 1).Lvl Then
                    Temp = GuildsInfo(b)
                    GuildsInfo(b) = GuildsInfo(b + 1)
                    GuildsInfo(b + 1) = Temp
                
                  ElseIf .Lvl = GuildsInfo(b + 1).Lvl Then
                    If .Exp < GuildsInfo(b + 1).Exp Then
                        Temp = GuildsInfo(b)
                        GuildsInfo(b) = GuildsInfo(b + 1)
                        GuildsInfo(b + 1) = Temp
                    End If
                End If
            End With
        Next b
    Next A
                
End Sub
Public Sub WriteGuilds_Found(ByVal Name As String, ByVal Alineation As eGuildAlineation)
    
    Dim A As Long
    
    Call Writer.WriteInt(ClientPacketID.Guilds_Found)
    Call Writer.WriteString8(Name)
    Call Writer.WriteInt8(Alineation)
    Call modNetwork.Send(False)
    
End Sub

Public Sub WriteGuilds_Invitation(ByVal UserName As String, ByVal Tipo As Byte)
    
    Call Writer.WriteInt(ClientPacketID.Guilds_Invitation)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt(Tipo)
    Call modNetwork.Send(False)
    
End Sub

Public Function SearchGuildSlot(ByVal GuildIndex As Integer) As Integer
    Dim A As Long
    
    For A = 1 To MAX_GUILDS
        If GuildsInfo(A).Index = GuildIndex Then
            SearchGuildSlot = A
            Exit Function
        End If
    Next A
End Function
    
Private Sub HandleGuild_Info()
    
    Dim List() As String
    Dim A      As Long
    Dim GuildIndex As Integer
    Dim Alineation As String, Colour As Long
    Dim Slot As Integer
    
    Call InitGrh(Grh_Antorcha, 25)
    GuildIndex = Reader.ReadInt16
    Slot = SearchGuildSlot(GuildIndex)

    With GuildsInfo(Slot)
        .Name = Reader.ReadString8
        .Alineation = Reader.ReadInt8
        .Colour = ARGB(197, 150, 0, 255) 'Guilds_Alineation_Colour(GuildsInfo(A).Alineation)
        
        For A = 1 To MAX_GUILD_MEMBER

            With .Members(A)
                .Name = Reader.ReadString8
                .Range = Reader.ReadInt8
                
                .Body = Reader.ReadInt16
                .Head = Reader.ReadInt16
                .Helm = Reader.ReadInt16
                .Shield = Reader.ReadInt16
                .Weapon = Reader.ReadInt16
            End With
        Next A

    End With
    
End Sub

Public Sub WriteGuilds_Online()

    Call Writer.WriteInt(ClientPacketID.Guilds_Online)
    Call modNetwork.Send(False)
End Sub

Private Sub HandleGuild_InfoUsers()
    
    Dim List() As String
    Dim A      As Long
    Dim GuildIndex As Integer
    
    'Call InitGrh(Grh_Antorcha, 25)
    
    GuildIndex = Reader.ReadInt16
    
    With GuildsInfo(GuildIndex)

        For A = 1 To MAX_GUILD_MEMBER

            With .Members(A)
                .Name = Reader.ReadString8
                .Range = Reader.ReadInt
                    
                .Elv = Reader.ReadInt
                .Class = Reader.ReadInt
                .Raze = Reader.ReadInt
                    
                .Body = Reader.ReadInt
                .Head = Reader.ReadInt
                .Helm = Reader.ReadInt
                .Shield = Reader.ReadInt
                .Weapon = Reader.ReadInt
                    
                .Points = Reader.ReadInt
            End With
        Next A

    End With
    
    Selected_GuildIndex = GuildIndex
    
    #If ModoBig = 1 Then
        dockForm FrmGuilds_Leader.hWnd, FrmMain.PicMenu, True
    #Else
        Call FrmGuilds_Leader.Show(, FrmMain)
    #End If


End Sub

Public Sub WriteGuilds_Kick(ByVal UserName As String)
    
    Call Writer.WriteInt(ClientPacketID.Guilds_Kick)
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)
    
End Sub

Public Sub WriteGuilds_Abandonate()

    Call Writer.WriteInt(ClientPacketID.Guilds_Abandonate)
    Call modNetwork.Send(False)
End Sub

Private Sub HandleFight_PanelInvitation()
    
    Dim List()                            As String

    Dim A                                 As Long

    Dim Text                              As String

    Dim TextFight                         As String

    Dim Temp                              As tFight

    Dim TempTeam(1)                       As String
    
    Dim Config(0 To MAX_RETOS_CONFIG - 1) As Byte
    
    Fight_UserName = Reader.ReadString8
    TextFight = Reader.ReadString8
    Temp.Gld = Reader.ReadInt32

    Temp.Rounds = Reader.ReadInt8
    Temp.Zona = Reader.ReadInt8
    
    Temp.LimiteTiempo = Reader.ReadInt8
    List = Split(TextFight, SEPARATOR)

    For A = 0 To MAX_RETOS_CONFIG - 1
        Config(A) = Reader.ReadInt8
    Next A
    
    If FightOn Then

        For A = LBound(List) To UBound(List)

            If A <= (UBound(List) / 2) Then
                TempTeam(0) = TempTeam(0) & List(A) & vbCrLf
            Else
                TempTeam(1) = TempTeam(1) & List(A) & vbCrLf

            End If

        Next A
    
        TempTeam(0) = Left$(TempTeam(0), Len(TempTeam(0)) - Len(vbCrLf))
        TempTeam(1) = Left$(TempTeam(1), Len(TempTeam(1)) - Len(vbCrLf))
    
        With FrmRetos

            For A = 0 To MAX_RETOS_CONFIG - 1
                FrmRetos.chkConfig(A).Enabled = False
                FrmRetos.chkConfig(A).Value = Config(A)
            Next A

            .txtGld.Text = Temp.Gld
            
            
            .cmbRounds.Text = Temp.Rounds
            .cmbTime.Text = Temp.LimiteTiempo

            
            For A = LBound(List) To UBound(List)
                List(A) = Replace$(List(A), vbCrLf, vbNullString)
                .txtUser(A).Text = List(A)
                .txtUser(A).Locked = True
                .txtUser(A).visible = True
            Next A
            
            .TypeFight = eTypeFight.eAccept
            
            #If ModoBig = 1 Then
                dockForm .hWnd, FrmMain.PicMenu, True
            #Else
                Call .Show(, FrmMain)
            #End If

        End With
    
    Else
        
        For A = LBound(List) To UBound(List)
            
            List(A) = Replace$(List(A), vbCrLf, vbNullString)
            
            If A <= (UBound(List) / 2) Then
                TempTeam(0) = TempTeam(0) & List(A) & ", "
            Else
                TempTeam(1) = TempTeam(1) & List(A) & ", "

            End If

        Next A
    
        TempTeam(0) = Left$(TempTeam(0), Len(TempTeam(0)) - 2)
        TempTeam(1) = Left$(TempTeam(1), Len(TempTeam(1)) - 2)
        Call ShowConsoleMsg(IIf((Temp.Tipo = 4), "Plante» ", vbNullString) & TempTeam(0) & " vs " & TempTeam(1) & vbCrLf & "Monedas de Oro: " & Temp.Gld & vbCrLf & "Utiliza /RETOSON para ver las solicitudes de retos de una forma más dinámica. Tipea /ACEPTAR " & Fight_UserName, 144, 251, 115)

        If Config(eRetoConfig.eItems) = 1 Then
        
            Call ShowConsoleMsg("¡Podrías perder todos tus objetos! ¡No retes si no estás seguro!", 255, 20, 20, True)

        End If
    
    End If
    
End Sub

Public Sub WriteFight_CancelInvitation()

    Call Writer.WriteInt(ClientPacketID.Fight_CancelInvitation)
    Call modNetwork.Send(False)
End Sub



Public Sub WriteGuilds_Talk(ByVal chat As String, ByVal Support As Boolean)
    
    Call Writer.WriteInt(ClientPacketID.Guilds_Talk)
        
    Call Writer.WriteString8(chat)
    Call Writer.WriteBool(Support)
    Call modNetwork.Send(False)

End Sub

Public Sub WriteLoginAccount()
    
    Call Writer.WriteInt(ClientPacketID.LoginAccount)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    
    'Call Writer.WriteString8(mEncrypt_A.AesEncryptString(Account.Email, mEncrypt_B.XOR_CHARACTER))
    'Call Writer.WriteString8(mEncrypt_A.AesEncryptString(Account.Passwd, mEncrypt_B.XOR_CHARACTER))
    Call Writer.WriteString8(Account.Email)
    Call Writer.WriteString8(Account.Passwd)
    
    Call Writer.WriteString8(AccountSec.SERIAL_BIOS)
    Call Writer.WriteString8(AccountSec.SERIAL_DISK)
    Call Writer.WriteString8(AccountSec.SERIAL_MAC)
    Call Writer.WriteString8(AccountSec.SERIAL_MOTHERBOARD)
    Call Writer.WriteString8(AccountSec.SERIAL_PROCESSOR)
    Call Writer.WriteString8(AccountSec.SYSTEM_DATA)
    Call Writer.WriteString8(AccountSec.IP_Local)
    Call Writer.WriteString8(AccountSec.IP_Public)
    
    Call modNetwork.Send(False)

End Sub
Public Sub WriteLoginAccountNew()
    
    'Call Writer.WriteInt(ClientPacketID.LoginAccountNew)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(Account.Email)
    
    Call modNetwork.Send(False)

End Sub

Public Sub WriteLoginChar()
    
    Call Writer.WriteInt(ClientPacketID.LoginChar)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(Account.Key)
    Call Writer.WriteInt8(Account.SelectedChar)
    
    Call modNetwork.Send(False)

End Sub

Public Sub WriteLoginCharNew()
    
    Call Writer.WriteInt(ClientPacketID.LoginCharNew)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt8(UserClase)
    Call Writer.WriteInt8(UserRaza)
    Call Writer.WriteInt8(UserSexo)
    Call Writer.WriteInt16(UserHead)

    Call modNetwork.Send(False)

End Sub

Public Sub WriteLoginName()
    
    Call Writer.WriteInt(ClientPacketID.LoginName)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(Account.Alias)
    
    Call modNetwork.Send(False)
    
End Sub

Public Sub WriteLoginRemove()
    
    Call Writer.WriteInt(ClientPacketID.LoginRemove)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(Account.Key)
    Call Writer.WriteInt8(Account.SelectedChar)
    
    Call modNetwork.Send(False)

End Sub
Public Sub WriteDisconnectForced()
    
    ' Writer.WriteInt(ClientPacketID.DisconnectForced)
    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(Account.Email)
    Call Writer.WriteString8(Account.Key)
    
    Call modNetwork.Send(False)

End Sub

Public Sub WriteLoginPasswd()
    
   ' Call Writer.WriteInt(ClientPacketID.LoginPasswd)

    Call Writer.WriteInt(App.Major)
    Call Writer.WriteInt(App.Minor)
    Call Writer.WriteInt(App.Revision)
    Call Writer.WriteString8(Account.Key)
    Call Writer.WriteString8(Account.Passwd)
    
    Call modNetwork.Send(False)

End Sub

Private Sub HandleLoggedAccount_DataChar()
    
    Dim SlotChar As Byte
    
    SlotChar = Reader.ReadInt8

    With Account.Chars(SlotChar)
        .Name = Reader.ReadString8
        .Guild = Reader.ReadString8
        
        .Body = Reader.ReadInt16
        .Head = Reader.ReadInt16
        .Weapon = Reader.ReadInt16
        .Shield = Reader.ReadInt16
        .Helm = Reader.ReadInt16
        
        .Ban = Reader.ReadInt8
        .Class = Reader.ReadInt8
        .Raze = Reader.ReadInt8
        .Elv = Reader.ReadInt8
        
        .PosMap = Reader.ReadInt16
        .PosX = Reader.ReadInt8
        .PosY = Reader.ReadInt8
        
        .Faction = Reader.ReadInt8
        .FactionRange = Reader.ReadInt8
    
        If .Faction = 1 Then
            .Colour = ARGB(50, 50, 255, 255)
        ElseIf .Faction = 2 Then
            .Colour = ARGB(255, 50, 50, 255)
        ElseIf .Faction = 3 Then
            .Colour = ARGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b, 255)
        ElseIf .Faction = 4 Then
            .Colour = ARGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b, 255)
        End If
    End With

End Sub

Private Sub HandleAccountInfo()

    Account.Gld = Reader.ReadInt32
    Account.Eldhir = Reader.ReadInt32
    Account.Premium = Reader.ReadInt8
    Account.SaleSlot = Reader.ReadInt16
    UserPoints = Reader.ReadInt32
    
    If FrmShop.visible Then
        FrmShop.lblDsp.Caption = PonerPuntos(Account.Eldhir)
        FrmShop.lblGld.Caption = PonerPuntos(Account.Gld)
        FrmShop.lblPoints.Caption = PonerPuntos(UserPoints)
    End If
End Sub

Private Sub HandleLoggedAccount()

        '<EhHeader>
        On Error GoTo HandleLoggedAccount_Err

        '</EhHeader>
    
        Dim A    As Long

        Dim Temp As String

        Dim Chars
100     EngineRun = True

102     '  Call SaveNewAccount(LCase$(Account.Email), Account.Passwd, False)
    
104   Account.Gld = Reader.ReadInt32
        Account.Eldhir = Reader.ReadInt32
        UserPoints = Reader.ReadInt32
        Account.Premium = Reader.ReadInt8
        Account.SaleSlot = Reader.ReadInt16
        Account.CharsAmount = Reader.ReadInt8

            
106     For A = 1 To ACCOUNT_MAX_CHARS
              Account.Chars(A).ID = Reader.ReadInt8
108         Account.Chars(A).Name = Reader.ReadString8
              Account.Chars(A).Blocked = Reader.ReadInt8
110         Account.Chars(A).Guild = Reader.ReadString8
        
112         Account.Chars(A).Body = Reader.ReadInt16
114         Account.Chars(A).Head = Reader.ReadInt16
116         Account.Chars(A).Weapon = Reader.ReadInt16
118         Account.Chars(A).Shield = Reader.ReadInt16
120         Account.Chars(A).Helm = Reader.ReadInt16
        
122         Account.Chars(A).Ban = Reader.ReadInt8
124         Account.Chars(A).Class = Reader.ReadInt8
126         Account.Chars(A).Raze = Reader.ReadInt8
128         Account.Chars(A).Elv = Reader.ReadInt8
        
130         Account.Chars(A).PosMap = Reader.ReadInt16
132         Account.Chars(A).PosX = Reader.ReadInt8
134         Account.Chars(A).PosY = Reader.ReadInt8
        
136         Account.Chars(A).Faction = Reader.ReadInt8
138         Account.Chars(A).FactionRange = Reader.ReadInt8

140         If Account.Chars(A).Faction = 1 Then
142             Account.Chars(A).Colour = ARGB(50, 50, 255, 255)
144         ElseIf Account.Chars(A).Faction = 2 Then
146             Account.Chars(A).Colour = ARGB(255, 50, 50, 255)
148         ElseIf Account.Chars(A).Faction = 3 Then
150             Account.Chars(A).Colour = ARGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b, 255)
152         ElseIf Account.Chars(A).Faction = 4 Then
154             Account.Chars(A).Colour = ARGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b, 255)

            End If

156     Next A
        
        If Not FrmMain.visible Then
            MirandoCuenta = True
   
160       FrmConnect_Account.Show
            FrmConnect_Account.SelectedPanelAccount (ePanelAccount)

        End If

164     'Unload FrmConnect
        '<EhFooter>
        Exit Sub

HandleLoggedAccount_Err:
        LogError err.Description & vbCrLf & "in ARGENTUM.Protocol.HandleLoggedAccount " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub




Private Sub HandleConnectedMessage()
    
    Reader.ReadInt8
    
    Call Login
    
End Sub

Private Sub HandleRender_CountDown()

    CountDownTime = Reader.ReadInt
    
    If CountDownTime = 0 Then
        CountDownTime_Fight = True
        CountDownTime = 255
    Else
        CountDownTime_Fight = False
    End If
    
End Sub


Private Sub HandleLoggedRemoveChar()
    Dim SlotUsername As Byte
    Dim CharAccount As tAccountChar
    
    SlotUsername = Reader.ReadInt
    
    Account.Chars(Account.SelectedChar) = CharAccount

    Account.SelectedChar = 0
    Account.CharsAmount = Account.CharsAmount - 1
    
   ' SwitchMap (1)
End Sub


''
' Writes the "BankGold" message to the outgoing data Reader.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankGold(ByVal Amount As Long, ByVal TypeGLD As Byte, ByVal Extract As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankGold" message to the outgoing data buffer
    '***************************************************
   
    Call Writer.WriteInt(ClientPacketID.BankGold)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt(TypeGLD)
    Call Writer.WriteBool(Extract)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteMercader_New(ByVal IsOffer As Integer, ByRef MU As tMercader)
    
    Call Writer.WriteInt(ClientPacketID.Mercader_New)
    
    Call Writer.WriteString8(Account.Passwd)
    Call Writer.WriteString8(Account.Key)
    
    ' Si es una oferta, envia el SLOT DE publicación al que deseamos ofertar
    Call Writer.WriteInt16(IsOffer)
    
    ' Envia la Nueva Publicacion/Oferta
    Call Writer.WriteInt32(MU.Gld)
    Call Writer.WriteInt32(MU.Dsp)
    Call Writer.WriteString8(MU.Desc)
    Call Writer.WriteInt32(MercaderGld)
    Call Writer.WriteInt8(MU.Blocked)
    Call Writer.WriteSafeArrayInt8(MU.IDCHARS)

    
    Call modNetwork.Send(False)

End Sub
Public Sub WriteMercader_Required(ByVal Required As Byte, _
                                  ByVal aBound As Integer, _
                                  ByVal bBound As Integer)
    
    Call Writer.WriteInt(ClientPacketID.Mercader_Required)
    Call Writer.WriteInt(Required)
    Call Writer.WriteInt(aBound)
    Call Writer.WriteInt(bBound)
    
    Call modNetwork.Send(False)

End Sub

Private Sub HandleMercader_List()

    Dim CharAmount   As Byte

    Dim A            As Long, b As Long

    Dim Text         As String

    Dim Amount       As Byte

    Dim Temp         As String

    Dim TempA()      As String
    
    Dim Index        As Integer
    
    Dim aBound       As Integer

    Dim bBound       As Integer
    
    Dim NullMercader As tMercader
    
    Dim Copy() As tMercader
    
    aBound = Reader.ReadInt16
    bBound = Reader.ReadInt16
    
    MercaderUserSlot = Reader.ReadInt16
    
    ReDim Copy(aBound To bBound) As tMercader
        
    ' El pibe cliclea por primer vez y solicita la lista
    If aBound = 1 Then
        
        MercaderID_Selected = 1
        MercaderLoaded = False

        For A = aBound To bBound
            Copy(A) = NullMercader
        Next A
        

    End If
    
    For A = aBound To bBound

        With Copy(A)
            .Loaded = True
            .ID = Reader.ReadInt16
            .Char = Reader.ReadInt8
            .Desc = Reader.ReadString8
            .Dsp = Reader.ReadInt32
            .Gld = Reader.ReadInt32
            
            For b = 1 To .Char
                .Chars(b).Name = Reader.ReadString8
                .Chars(b).Class = Reader.ReadInt8
                .Chars(b).Raze = Reader.ReadInt8

                .Chars(b).Elv = Reader.ReadInt8
                .Chars(b).Exp = Reader.ReadInt32
                .Chars(b).Elu = Reader.ReadInt32
                    
                .Chars(b).Hp = Reader.ReadInt16
                .Chars(b).Constitucion = Reader.ReadInt8
                
                .Chars(b).Body = Reader.ReadInt16
                .Chars(b).Head = Reader.ReadInt16
                .Chars(b).Weapon = Reader.ReadInt16
                .Chars(b).Shield = Reader.ReadInt16
                .Chars(b).Helm = Reader.ReadInt16
        
                If .Chars(b).Elv > 0 Then
                    Call Mercader_GenerateText(.Chars(b))
                End If
               
            Next b
            
        End With

    Next A
    
    If MercaderOff = 3 Then
        Mercader_ModoOferta = True
        MercaderOff = 0
        
        For A = LBound(Copy) To UBound(Copy)
            MercaderListOffer(A) = Copy(A)
        Next A

        If FrmMercaderOffers.visible Then
            FrmMercaderOffers.UpdateInfo
        Else
            FrmMercaderOffers.Show , FrmMain
        End If
        
        
    Else
    
        For A = LBound(Copy) To UBound(Copy)
            MercaderList_Copy(A) = Copy(A)
        Next A
        
        For A = LBound(Copy) To UBound(Copy)
            MercaderList(A) = MercaderList_Copy(A)
        Next A
        If Not FrmMercaderList.visible Then
            FrmMercaderList.Show , FrmMain
        End If
    End If
    

End Sub

Private Sub HandleMercader_ListChar()

    Dim SlotChar As Integer

    Dim A        As Long
    
    SlotChar = Reader.ReadInt16
    
    With MercaderList_Copy(MercaderID).Chars(SlotChar)
        .Name = Reader.ReadString8
        .Guild = Reader.ReadString8
                
        .Gld = Reader.ReadInt32
                
        .Body = Reader.ReadInt16
        .Head = Reader.ReadInt16
        .Weapon = Reader.ReadInt16
        .Shield = Reader.ReadInt16
        .Helm = Reader.ReadInt16
                
        .Faction = Reader.ReadInt8
        .FactionRange = Reader.ReadInt8
        .FragsCiu = Reader.ReadInt16
        .FragsCri = Reader.ReadInt16
        
        
        For A = 1 To MAX_INVENTORY_SLOTS
            .Object(A).ObjIndex = Reader.ReadInt16
            .Object(A).Amount = Reader.ReadInt16
        Next
        
        For A = 1 To MAX_BANCOINVENTORY_SLOTS
            .Bank(A).ObjIndex = Reader.ReadInt16
            .Bank(A).Amount = Reader.ReadInt16
        Next A
        
        For A = 1 To 35
            .Spells(A) = Reader.ReadString8
            
            If .Spells(A) <> vbNullString Then
                hlstMercader.AddItem (.Spells(A))
            End If
        Next A
        
        For A = 1 To NUMSKILLS
            .Skills(A) = Reader.ReadInt8
        Next A
        
        

    End With
    
End Sub

Private Sub HandleMercader_ListOffer()
    Dim CharAmount As Byte
    Dim A As Long, b As Long
    Dim Text As String
    Dim Amount As Byte
    Dim Temp As String
    

End Sub


Public Sub WriteForgive_Faction()
    Call Writer.WriteInt(ClientPacketID.Forgive_Faction)

    Call modNetwork.Send(False)

End Sub

Public Sub WriteMap_RequiredInfo(ByVal Map As Integer)
    Call Writer.WriteInt(ClientPacketID.Map_RequiredInfo)
    Call Writer.WriteInt(Map)
    
    Call modNetwork.Send(False)
End Sub

Private Sub HandleMiniMap_InfoCriature()

    
    Dim Map As Long
    Dim NpcsNum     As Byte
    Dim A           As Long, b As Long
    Dim Npcs As Integer
    
    Map = Reader.ReadInt
    NpcsNum = Reader.ReadInt

    MapSelected = Map

    MiniMap(Map).Name = Reader.ReadString8
    MiniMap(Map).Pk = Reader.ReadBool
    MiniMap(Map).LvlMin = Reader.ReadInt
    MiniMap(Map).LvlMax = Reader.ReadInt
    MiniMap(Map).Loaded = True
    
    'frmMapa.lblMap.Caption = "Mapa " & MapSelected & " (" & MiniMap(MapSelected).Name & ")"

    If NpcsNum Then
        MiniMap(Map).NpcsNum = NpcsNum
        
        For A = 1 To NpcsNum
            With MiniMap(Map).Npcs(A)
                .NpcIndex = Reader.ReadInt16
                .Name = Reader.ReadString8
                .Body = Reader.ReadInt
                .Head = Reader.ReadInt
                .Hp = Reader.ReadInt
                .MinHit = Reader.ReadInt
                .MaxHit = Reader.ReadInt
                .Exp = Reader.ReadInt
                .Gld = Reader.ReadInt
                .Eldhir = Reader.ReadInt
                
                .NroSpells = Reader.ReadInt
                
                If .NroSpells Then
                    ReDim .Spells(1 To .NroSpells) As String
                    
                    For b = 1 To .NroSpells
                        .Spells(b) = Reader.ReadString8
                    Next b
                End If
                
                .NroItems = Reader.ReadInt
                
                For b = 1 To .NroItems
                    .Obj(b).Name = Reader.ReadString8
                    .Obj(b).Amount = Reader.ReadInt
                Next b
                
                .NroDrops = Reader.ReadInt
                
                For b = 1 To .NroDrops
                    .Drop(b).Name = Reader.ReadString8
                    .Drop(b).Amount = Reader.ReadInt
                    .Drop(b).Probability = Reader.ReadInt
                Next b
            End With
        Next A
    End If
    
End Sub

Public Sub WriteWherePower()
    Call Writer.WriteInt(ClientPacketID.WherePower)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteAuction_New(ByVal Slot As Byte, ByVal Amount As Integer, ByVal Gld As Long, ByVal Eldhir As Long)
    Call Writer.WriteInt(ClientPacketID.Auction_New)
        
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt(Gld)
    Call Writer.WriteInt(Eldhir)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteAuction_Info()
    Call Writer.WriteInt(ClientPacketID.Auction_Info)
    Call modNetwork.Send(False)
End Sub
Public Sub WriteAuction_Offer(ByVal Gld As Long, ByVal Eldhir As Long)
    Call Writer.WriteInt(ClientPacketID.Auction_Offer)
    Call Writer.WriteInt(Gld)
    Call Writer.WriteInt(Eldhir)
    Call modNetwork.Send(False)
End Sub
Public Sub WriteGoInvation(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.GoInvation)
    Call Writer.WriteInt8(Slot)
    Call modNetwork.Send(False)
End Sub

Public Sub Crafting_ClassValid(ByVal SlotObjIndex As Integer, _
                                ByRef ClassInValid() As Byte)
    Dim A As Long, b As Long
    Dim Temp As String
    Dim Count As Byte
    Dim Valid As Byte
    
    With ObjBlacksmith_Copy(SlotObjIndex)
        For A = 1 To NUMCLASES
            .ClassValid(A) = 1
            
            For b = LBound(ClassInValid) To UBound(ClassInValid)
                If A = ClassInValid(b) Then
                    .ClassValid(A) = 0
                End If
            Next b
            
            If .ClassValid(A) = 1 Then
                Count = Count + 1
            End If
        Next A
        
        If Count = NUMCLASES Then
            .ValidTotal = True
        End If
    
    End With
End Sub
Public Sub WriteSendDataUser(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SendDataUser)
        
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)
End Sub

 #If ClienteGM = 1 Then
Public Sub WriteSearchDataUser(ByVal Selected As eSearchData, ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.SearchDataUser)
        
    Call Writer.WriteInt8(Selected)
    Call Writer.WriteString8(UserName)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteChangeModoArgentum()
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.ChangeModoArgentum)
        
    Call modNetwork.Send(False)
End Sub




#End If

Public Sub WriteStreamerBotSetting(ByVal Delay As Long, ByVal Mode As Byte, ByVal DelayIndex As Long)
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.StreamerBotSetting)
    Call Writer.WriteInt32(Delay)
    Call Writer.WriteInt8(Mode)
    Call Writer.WriteInt32(DelayIndex)
    
    Call modNetwork.Send(False)
End Sub
Private Sub HandleUpdateEffectPoison()
    
    UserEnvenenado = Not UserEnvenenado
    
End Sub


Public Sub WriteEvents_DonateObject(ByVal Slot As Byte, ByVal Amount As Integer)
    
    Call Writer.WriteInt(ClientPacketID.Events_DonateObject)
    Call Writer.WriteInt8(TEMP_SLOTEVENT)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
    
    Call modNetwork.Send(False)
End Sub

Private Sub HandleRenderConsole()
    
    Dim Text As String
    Dim DamageType As EDType
    Dim Duration As Long
    Dim Slot As Byte
    
    Text = Reader.ReadString8
    DamageType = Reader.ReadInt8
    Duration = Reader.ReadInt32
    Slot = Reader.ReadInt8
    
    Call RenderText_Console_Add(Text, DamageType, Duration, Slot)
    
End Sub

Private Sub HandleViewListQuest()
    
    Dim cant   As Byte, A As Long, b As Long
    Dim List() As Byte, DataTemp As tQuest
    
    cant = Reader.ReadInt8: QuestLast = cant
    NpcName = Reader.ReadString8
    
    ReDim QuestNpc(1 To cant) As Byte
    
    For A = 1 To cant
        QuestNpc(A) = Reader.ReadInt8
    Next A
    
    frmCriatura_Quest.Show , FrmMain
    Render_QuestPanel
End Sub

Private Sub HandleUpdateUserDead()
    CharList(UserCharIndex).Muerto = Reader.ReadInt8
    
   ' If CharList(UserCharIndex).Muerto Then
       ' Set_engineBaseSpeed (0.024)
    'Else
     '  Set_engineBaseSpeed (0.018)
    'End If
    

End Sub
Public Sub WriteQuestRequired(ByVal QuestRequired As Byte)
    
    Call Writer.WriteInt(ClientPacketID.QuestRequired)
    Call Writer.WriteInt8(QuestRequired)
    
    Call modNetwork.Send(False)
End Sub

' Chequea si los objetos se pueden visualizar. Segun la clase del personaje.
Public Sub Quests_CheckViewObjs(ByVal QuestIndex As Integer)

    Dim A As Long, b As Long
    
    With QuestList(QuestIndex)

        If .RewardObj > 0 Then
            
            For A = 1 To .RewardObj
                .RewardObjs(A).View = True
                
                If ObjData(.RewardObjs(A).ObjIndex).CP_Valid Then
                    
                    For b = LBound(ObjData(.RewardObjs(A).ObjIndex).CP) To UBound(ObjData(.RewardObjs(A).ObjIndex).CP)
                         
                         
                        If ObjData(.RewardObjs(A).ObjIndex).CP(b) = UserClase Then
                            .RewardObjs(A).View = False
                            Exit For
                        End If

                    Next b

                End If

            Next A
                
        End If
    
    End With

End Sub

Private Sub HandleQuestData()

        '<EhHeader>
        On Error GoTo HandleQuestData_Err

        '</EhHeader>
    
        Dim A       As Long, b As Long

        Dim MaxHp   As Long

        Dim visible As Boolean
    
        Dim Slot    As Integer

100     PuedeReclamar = True
    
108     visible = Reader.ReadBool
110     Slot = Reader.ReadInt16
        
        If visible Then
            Set FrmObjetive.ListQuests = New clsGraphicalList
            Call FrmObjetive.ListQuests.Initialize(FrmObjetive.picQuest, RGB(200, 190, 190), 14, 30)

        End If

112     If Slot > 0 Then
114         NpcsUser_QuestIndex = Reader.ReadInt16
116         NpcsUser_QuestIndex_Original = NpcsUser_QuestIndex

118         If NpcsUser_QuestIndex > 0 Then
            
120             With QuestList(NpcsUser_QuestIndex)
122
124                 .QuestEmpezada = True
126                 NpcsUser_Selected = 1
        
128                 For A = 1 To .Npc
130                     MaxHp = NpcList(.Npcs(A).NpcIndex).MaxHp
132                     .NpcsUser(A).Amount = Reader.ReadInt32
134                     '.NpcsUser(A).Amount = Int((.NpcsUser(A).Amount) / MaxHp)
                  
136                     If (.NpcsUser(A).Amount / NpcList(.Npcs(A).NpcIndex).MaxHp) <> .Npcs(A).Amount Then
138                         PuedeReclamar = False
140                         .NpcsUser(A).Color = ARGB(130, 130, 130, 255)
                        Else
142                         .NpcsUser(A).Color = ARGB(38, 137, 16, 255)

                        End If

144                 Next A
              
146                 For A = 1 To .SaleObj
148                     .ObjsSaleUser(A).Amount = Reader.ReadInt32

150                     If .ObjsSaleUser(A).Amount <> .SaleObjs(A).Amount Then
152                         PuedeReclamar = False
154                         .ObjsSaleUser(A).Color = ARGB(130, 130, 130, 255)
                        Else
156                         .ObjsSaleUser(A).Color = ARGB(16, 137, 118, 255)

                        End If

158                 Next A
        
160                 For A = 1 To .ChestObj
162                     .ObjsChestUser(A).Amount = Reader.ReadInt32

164                     If .ObjsChestUser(A).Amount <> .ChestObjs(A).Amount Then
166                         PuedeReclamar = False
168                         .ObjsChestUser(A).Color = ARGB(130, 130, 130, 255)
                        Else
170                         .ObjsChestUser(A).Color = ARGB(150, 80, 130, 255)

                        End If

172                 Next A
        
174                 For A = 1 To .Obj
176                     .ObjsUser(A).Amount = TieneObjetos(.Objs(A).ObjIndex)

178                     If .ObjsUser(A).Amount > .Objs(A).Amount Then .ObjsUser(A).Amount = .Objs(A).Amount
            
180                     If .ObjsUser(A).Amount <> .Objs(A).Amount Then
182                         PuedeReclamar = False
184                         .ObjsUser(A).Color = ARGB(130, 130, 130, 255)
                        Else
            
186                         .ObjsUser(A).Color = ARGB(222, 140, 20, 255)

                        End If

188                 Next A
            
                End With
        
190             Call Quests_CheckViewObjs(NpcsUser_QuestIndex)
         
            Else
192             FrmObjetive.ListQuests.AddItem "(Disponible)"

            End If
    
        Else
            FrmObjetive.ListQuests.Clear
              
194         For b = 1 To MAXUSERQUESTS
                  
196             NpcsUser_Quest(b) = Reader.ReadInt16

198             If NpcsUser_QuestIndex = 0 Then
200                 NpcsUser_QuestIndex = NpcsUser_Quest(b)
202                 NpcsUser_QuestIndex_Original = NpcsUser_QuestIndex

                End If

204             If NpcsUser_Quest(b) > 0 Then
                        
206                 With QuestList(NpcsUser_Quest(b))
208                     FrmObjetive.ListQuests.AddItem .Name
210                     .QuestEmpezada = True
212                     NpcsUser_Selected = 1
        
214                     For A = 1 To .Npc
216                         MaxHp = NpcList(.Npcs(A).NpcIndex).MaxHp
218                         .NpcsUser(A).Amount = Reader.ReadInt32
220                         '.NpcsUser(A).Amount = Int((.NpcsUser(A).Amount) / MaxHp)
                  
222                         If (.NpcsUser(A).Amount / NpcList(.Npcs(A).NpcIndex).MaxHp) <> .Npcs(A).Amount Then
224                             PuedeReclamar = False
226                             .NpcsUser(A).Color = ARGB(130, 130, 130, 255)
                            Else
228                             .NpcsUser(A).Color = ARGB(38, 137, 16, 255)

                            End If

230                     Next A
              
232                     For A = 1 To .SaleObj
234                         .ObjsSaleUser(A).Amount = Reader.ReadInt32

236                         If .ObjsSaleUser(A).Amount <> .SaleObjs(A).Amount Then
238                             PuedeReclamar = False
240                             .ObjsSaleUser(A).Color = ARGB(130, 130, 130, 255)
                            Else
242                             .ObjsSaleUser(A).Color = ARGB(16, 137, 118, 255)

                            End If

244                     Next A
        
246                     For A = 1 To .ChestObj
248                         .ObjsChestUser(A).Amount = Reader.ReadInt32

250                         If .ObjsChestUser(A).Amount <> .ChestObjs(A).Amount Then
252                             PuedeReclamar = False
254                             .ObjsChestUser(A).Color = ARGB(130, 130, 130, 255)
                            Else
256                             .ObjsChestUser(A).Color = ARGB(150, 80, 130, 255)

                            End If

258                     Next A
        
260                     For A = 1 To .Obj
262                         .ObjsUser(A).Amount = TieneObjetos(.Objs(A).ObjIndex)

264                         If .ObjsUser(A).Amount >= .Objs(A).Amount Then .ObjsUser(A).Amount = .Objs(A).Amount
            
266                         If .ObjsUser(A).Amount <> .Objs(A).Amount Then
268                             PuedeReclamar = False
270                             .ObjsUser(A).Color = ARGB(130, 130, 130, 255)
                            Else
            
272                             .ObjsUser(A).Color = ARGB(222, 140, 20, 255)

                            End If

274                     Next A
            
                    End With
        
276                 Call Quests_CheckViewObjs(NpcsUser_Quest(b))
         
                Else
278                 FrmObjetive.ListQuests.AddItem "(Disponible)"

                End If

280         Next b
    
        End If
    
282     If visible Then
            #If ModoBig = 1 Then
284             dockForm FrmObjetive.hWnd, FrmMain.PicMenu, True
            #Else
    
286             If Not FrmObjetive.visible Then
288                 FrmObjetive.Show , FrmMain
    
                End If
    
            #End If

        End If

        '<EhFooter>
        Exit Sub

HandleQuestData_Err:
        LogError err.Description & vbCrLf & "in ARGENTUM.Protocol.HandleQuestData " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

Private Sub HandleUpdateGlobalCounter()

    Dim Tipo    As Byte
    Dim Counter As Integer
    
    Tipo = Reader.ReadInt8
    Counter = Reader.ReadInt16
    
    Select Case Tipo
    
        Case 1 ' Invisiblidad
            FrmMain.imgInvisible.visible = IIf(Counter > 0, True, False)
            FrmMain.lblInvi.visible = IIf(Counter > 0, True, False)
            FrmMain.lblInvi.Caption = Counter
        Case 2 ' Paralisis
            FrmMain.imgParalisis.visible = IIf(Counter > 0, True, False)
            FrmMain.lblParalisis.visible = IIf(Counter > 0, True, False)
            FrmMain.lblParalisis = Counter
            
        Case 3 ' Tiempo de la Dopardi
            FrmMain.imgDopa.visible = IIf(Counter > 0, True, False)
            FrmMain.lblDopa.visible = IIf(Counter > 0, True, False)
            FrmMain.lblDopa = Counter
        Case 4 ' Tiempo de regreso a Hogar
            FrmMain.imgHome.visible = IIf(Counter > 0, True, False)
            FrmMain.lblHome.visible = IIf(Counter > 0, True, False)
            FrmMain.lblHome = Counter
            
    End Select
    
End Sub

Private Sub HandleSendInfoNpc()
    Dim NpcIndex As Integer
    NpcIndex = Reader.ReadInt16
    
    SelectedNpcIndex = NpcIndex
    
    If Not FrmCriatura_Info.visible Then
        Call FrmCriatura_Info.Show(, FrmMain)
    End If
    
End Sub

Private Sub HandleUpdatePosGuild()
    
    Dim SlotMember As Byte
    Dim X As Byte
    Dim Y As Byte
    
    SlotMember = Reader.ReadInt8
    X = Reader.ReadInt8
    Y = Reader.ReadInt8
    
    MiniMap_Friends(SlotMember).X = X
    MiniMap_Friends(SlotMember).Y = Y
End Sub

Private Sub HandleUpdateLevelGuild()
    GuildLevel = Reader.ReadInt8
End Sub


Private Sub HandleReceiveIntervals()
    
    IntervaloUserPuedeAtacar = Reader.ReadInt16
    IntervaloUserPuedeUsar = Reader.ReadInt16
    IntervaloUserPuedeUsarClick = Reader.ReadInt16
    IntervaloUpdatePos = Reader.ReadInt16
    IntervaloUserPuedeCastear = Reader.ReadInt16
    IntervaloUserPuedeShiftear = Reader.ReadInt16
    IntervaloFlechasCazadores = Reader.ReadInt16
    IntervaloMagiaGolpe = Reader.ReadInt16
    IntervaloGolpeMagia = Reader.ReadInt16
    IntervaloGolpeUsar = Reader.ReadInt16
    IntervaloUserPuedeTrabajar = Reader.ReadInt16
    IntervalDrop = Reader.ReadInt16
    IntervaloCaminar = Reader.ReadReal32

    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, IntervaloUserPuedeAtacar)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, IntervaloUserPuedeUsar)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, IntervaloUserPuedeUsarClick)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, IntervaloUpdatePos)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, IntervaloUserPuedeCastear)
    Call MainTimer.SetInterval(TimersIndex.Shift, IntervaloUserPuedeShiftear - 150)
    Call MainTimer.SetInterval(TimersIndex.Arrows, IntervaloFlechasCazadores)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, IntervaloMagiaGolpe)
    Call MainTimer.SetInterval(TimersIndex.AttackSpell, IntervaloGolpeMagia)
    Call MainTimer.SetInterval(TimersIndex.AttackUse, IntervaloGolpeUsar)
    Call MainTimer.SetInterval(TimersIndex.Drop, IntervalDrop)
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar)
    Call MainTimer.SetInterval(TimersIndex.Work, IntervaloUserPuedeTrabajar)
    
    Call MainTimer.SetInterval(TimersIndex.Packet250, 250)
    Call MainTimer.SetInterval(TimersIndex.Packet500, 500)
    
    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Shift)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.AttackSpell)
    Call MainTimer.Start(TimersIndex.AttackUse)
    Call MainTimer.Start(TimersIndex.Drop)
    Call MainTimer.Start(TimersIndex.Walk)
    Call MainTimer.Start(TimersIndex.Packet250)
    Call MainTimer.Start(TimersIndex.Packet500)
    Call MainTimer.Start(TimersIndex.Work)
    
    'frmMain.MacroTrabajo.Interval = IntervaloUserPuedeTrabajar
End Sub

Private Sub HandleUpdateStatusMAO()
    Dim Status As Byte
    
    Status = Reader.ReadInt8
    
    Select Case Status
    
        Case 1 ' Nos ponemos en espera de codigo de confirmacion
            EsperandoValidacion = Not EsperandoValidacion
        
        Case 2 ' Unload FrmMercader
        
        Case 3 ' Unload FrmMercader_List
    End Select
End Sub



Private Sub HandleUpdateOnline()
    UsuariosOnline = Reader.ReadInt16
    
    FrmMain.lblOns.Caption = UsuariosOnline
End Sub

Private Sub HandleModoEvento()
    EsModoEvento = Reader.ReadInt8
    
End Sub

Public Sub WriteModoStream()
    Call Writer.WriteInt(ClientPacketID.ModoStreamer)

    Call modNetwork.Send(False)
End Sub


Public Sub WriteStreamerLink(ByVal Url As String)
    Call Writer.WriteInt(ClientPacketID.StreamerSetLink)
    Call Writer.WriteString8(Url)
    
    Call modNetwork.Send(False)
End Sub

Public Sub WriteChangeNick(ByVal UserName As String, _
                                         ByVal Leader As Boolean)
                                         
    Call Writer.WriteInt(ClientPacketID.ChangeNick)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteBool(Leader)
    
    Call modNetwork.Send(False)
End Sub

Public Sub WriteConfirmTransaccion(ByVal Email As String, _
                                   ByVal Promotion As Byte, _
                                   ByVal Bank As String)
                                   
    Call Writer.WriteInt(ClientPacketID.ConfirmTransaccion)
    Call Writer.WriteString8(Email)
    Call Writer.WriteInt8(Promotion)
    Call Writer.WriteString8(Bank)
    
    Call modNetwork.Send(False)

End Sub

Public Sub WriteConfirmItem(ByVal Slot As Integer, ByVal ModoPrice As Byte)
    Call Writer.WriteInt(ClientPacketID.ConfirmItem)
    Call Writer.WriteInt16(Slot)
    Call Writer.WriteInt8(ModoPrice)
    
    Call modNetwork.Send(False)
End Sub
Public Sub WriteConfirmTier(ByVal Tier As Integer)
    Call Writer.WriteInt(ClientPacketID.ConfirmTier)
    Call Writer.WriteInt8(Tier)
    
    Call modNetwork.Send(False)
End Sub

Public Sub HandleUpdateMeditation()
    
    ClientMeditation = Reader.ReadInt16
    
    Dim A As Long
    
    Dim Last As Byte
    
    Last = Reader.ReadInt8
    
    ReDim MeditationUser(1 To Last) As Byte
    
    For A = 1 To Last
        MeditationUser(A) = Reader.ReadInt8
    Next A
    
    If MirandoShop Then
        Call FrmShop.UpdateMeditationLearn
    End If
End Sub



Public Sub WriteRequiredShopChars()
    Call Writer.WriteInt(ClientPacketID.RequiredShopChars)
    
    Call modNetwork.Send(False)
End Sub

Private Sub HandleSendShopChars()

    ShopCharLast = Reader.ReadInt8
    
    Dim A As Long
    
    
    ReDim ShopChars(0 To ShopCharLast) As tShopChars
    
    For A = 1 To ShopCharLast
        With ShopChars(A)
            .Name = Reader.ReadString8
            .Dsp = Reader.ReadInt16
            .Elv = Reader.ReadInt8
            .Porc = Reader.ReadInt8
            .Class = Reader.ReadInt8
            .Raze = Reader.ReadInt8
            .Head = Reader.ReadInt16
            .Hp = Reader.ReadInt16
            .Man = Reader.ReadInt16
            
            If .Elv > 0 Then
                .Ups = UserCheckPromedy(.Elv, .Hp, .Class, ModRaza(.Raze).Constitucion)
            End If
        End With
    Next A
    
    
End Sub
Public Sub WriteConfirmChar(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.ConfirmChar)
    Call Writer.WriteInt8(Slot)
    
    Call modNetwork.Send(False)
End Sub

Public Sub WriteConfirmQuest(ByVal Tipo As Byte, ByVal Quest As Byte)
    Call Writer.WriteInt(ClientPacketID.ConfirmQuest)
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteInt8(Quest)
    
    Call modNetwork.Send(False)
End Sub

Private Sub HandleUpdateFinishQuest()

    Dim QuestIndex As Integer
    Dim List() As String
    Dim A As Long
    
    QuestIndex = Reader.ReadInt16
    
    ReDim Preserve List(0) As String

    With QuestList(QuestIndex)
        If .RewardExp > 0 Then
             ReDim Preserve List(0 To UBound(List) + 1) As String
             List(UBound(List)) = "+" & .RewardExp & " EXP"
        End If
        
        If .RewardGld > 0 Then
            ReDim Preserve List(0 To UBound(List) + 1) As String
             List(UBound(List)) = "+" & .RewardGld & " ORO"
        End If
        
        If .RewardObj > 0 Then
            For A = 1 To .RewardObj
                 ReDim Preserve List(0 To UBound(List) + 1) As String
                List(UBound(List)) = "+" & ObjData(.RewardObjs(A).ObjIndex).Name & " (x" & .RewardObjs(A).Amount & ")"
             Next A
        End If
        
        
        If Len(.DescFinish) Then
            Call ShowConsoleMsg(.DescFinish, 245, 212, 7)
        End If
        
    End With
    
    
    
    
    Call Anuncio_AddNew("Misión Completada", List, A_MISION_COMPLETADA)
    
End Sub
Public Sub WriteRequiredSkins(ByVal ObjIndex As Integer, ByVal Modo As Byte)
    Call Writer.WriteInt(ClientPacketID.RequiredSkins)

    Call Writer.WriteInt16(ObjIndex)
    
    Call Writer.WriteInt8(Modo)
    Call modNetwork.Send(False)
    
End Sub



Public Sub HandleUpdateDataSkin()
        '<EhHeader>
        On Error GoTo HandleUpdateDataSkin_Err
        '</EhHeader>

        Dim A As Long
        Dim ExistUser As Integer
        Dim Last As Integer
        
100     With ClientInfo.Skin
            Last = Reader.ReadInt16
            
104         If Last > 0 Then
               
               .Last = Last
               
106            ReDim .ObjIndex(1 To .Last) As Integer
            
108            For A = 1 To .Last
110                 .ObjIndex(A) = Reader.ReadInt16
                    
                    If Not InventorySkins Is Nothing Then
                    If InventorySkins.SelectedItem > 0 And .ObjIndex(A) > 0 Then
                        If InventorySkins.ObjIndex(InventorySkins.SelectedItem) = .ObjIndex(A) Then
                            If Skins_CheckingItems(.ObjIndex(A)) Then
                                FrmSkin.lblBuy.Caption = "DESEQUIPAR"
                            Else
                                FrmSkin.lblBuy.Caption = "USAR"
                            End If
                            
                            
                        End If
                    End If
                    End If

112             Next A
            End If
            
            .ArmourIndex = Reader.ReadInt16
            .HelmIndex = Reader.ReadInt16
            .ShieldIndex = Reader.ReadInt16
            .WeaponIndex = Reader.ReadInt16
            .WeaponArcoIndex = Reader.ReadInt16
            .WeaponDagaIndex = Reader.ReadInt16
            
            #If ModoBig = 1 Then
122             dockForm FrmSkin.hWnd, FrmMain.PicMenu, True
            #Else
124             FrmSkin.Show , FrmMain
            #End If
            
            Call FrmSkin.Skins_Load
        End With

        '<EhFooter>
        Exit Sub

HandleUpdateDataSkin_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Protocol.HandleUpdateDataSkin " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub HandleRequiredMoveChar()
    Dim Heading As Byte
    Dim LegalOk As Boolean
    
    Heading = Reader.ReadInt8
    
      Select Case Heading

        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)

        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)

        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)

        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If Not UserMoving Then
        Call MoveTo(Heading)
    End If
End Sub


Private Sub HandleUpdateBar()
    
    Dim CharIndex As Integer

    Dim BarMin     As Long

    Dim BarMax     As Long
    
    Dim Tipo As Byte
    
    Tipo = Reader.ReadInt8
    CharIndex = Reader.ReadInt16
    BarMin = Reader.ReadInt32
    BarMax = Reader.ReadInt32
    
    With CharList(CharIndex)
        .BarMin = BarMin
        .BarMax = BarMax
    End With

End Sub

Private Sub HandleUpdateBarTerrain()
    
    Dim X      As Integer

    Dim Y      As Integer
    
    Dim BarMin As Long

    Dim BarMax As Long

    Dim Tipo   As Byte
    
    Tipo = Reader.ReadInt8
    X = Reader.ReadInt16
    Y = Reader.ReadInt16
    BarMin = Reader.ReadInt32
    BarMax = Reader.ReadInt32
    
    With MapData(X, Y)
        .BarMin = BarMin
        .BarMax = BarMax
        
        If (.BarMin = 0 And .BarMax = 0) Then

            ' Borde de la Barra
            With GrhData(BAR_BORDER)
                Call g_Swarm.Remove(7, BAR_BORDER, 0, 0, 0, 0)

            End With
    
            ' Fondo de la Barra
            With GrhData(BAR_BACKGROUND)
                Call g_Swarm.Remove(7, BAR_BACKGROUND, 0, 0, 0, 0)

            End With

        Else

            ' Borde de la Barra
            With GrhData(BAR_BORDER)
                Call g_Swarm.Insert(7, BAR_BORDER, X, Y, .TileWidth, .TileHeight)

            End With
    
            ' Fondo de la Barra
            With GrhData(BAR_BACKGROUND)
                Call g_Swarm.Insert(7, BAR_BACKGROUND, X, Y, .TileWidth, .TileHeight)

            End With
        
        End If

    End With
End Sub


Public Sub WriteRequiredLive()
    
    Call Writer.WriteInt(ClientPacketID.RequiredLive)
    Call modNetwork.Send(False)
End Sub
Public Sub WriteAcelerationChar()
    
    Call Writer.WriteInt(ClientPacketID.AcelerationChar)
    Call modNetwork.Send(False)
End Sub
Private Sub HandleVelocidadToggle()
    
    On Error GoTo HandleVelocidadToggle_Err

    If UserCharIndex = 0 Then Exit Sub
    
    CharList(UserCharIndex).Speeding = Reader.ReadReal32()
    
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar / CharList(UserCharIndex).Speeding)
    
    Exit Sub

HandleVelocidadToggle_Err:
    
End Sub

Private Sub HandleSpeedToChar()
    
    On Error GoTo HandleSpeedToChar_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    
    Dim CharIndex As Integer

    Dim Speeding  As Single
     
    CharIndex = Reader.ReadInt16()
    Speeding = Reader.ReadReal32()
   
    CharList(CharIndex).Speeding = Speeding
    
    Exit Sub

HandleSpeedToChar_Err:
    
End Sub

Private Sub HandleUpdateUserTrabajo()
    
   If FrmMain.MacroTrabajo Then Call FrmMain.DesactivarMacroTrabajo
End Sub

Public Sub WriteAlquilar(ByVal Tipo As Byte)
    Call Writer.WriteInt(ClientPacketID.AlquilarComerciante)
    Call Writer.WriteInt8(Tipo)
    Call modNetwork.Send(False)
End Sub
Public Sub WriteTirarRuleta(ByVal Modo As Byte)
    Call Writer.WriteInt(ClientPacketID.TirarRuleta)
    Call Writer.WriteInt8(Modo)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteLotteryNew(ByRef TempLottery As tLottery)
    
    Call Writer.WriteInt(ClientPacketID.GmCommands)
    Call Writer.WriteInt(eGMCommands.LotteryNew)
    
    Call Writer.WriteString8(TempLottery.Name)
    Call Writer.WriteString8(TempLottery.Desc)
    
    Call Writer.WriteString8(TempLottery.DateFinish)
    
    Call Writer.WriteString8(TempLottery.PrizeChar)
    Call Writer.WriteInt16(TempLottery.PrizeObj)
    Call Writer.WriteInt16(TempLottery.PrizeObjAmount)
    
    Call modNetwork.Send(False)

End Sub
Private Sub HandleTournamentList()
    
    Dim A As Long, b As Long
    
    
    For A = 1 To MAX_EVENT_SIMULTANEO
        With Events(A)
            ReDim .AllowedClasses(1 To NUMCLASES) As Byte
            
            If Reader.ReadBool Then
                .Name = Reader.ReadString8
                .Config(eConfigEvent.eFuegoAmigo) = Reader.ReadInt8
                .LimitRound = Reader.ReadInt8
                .LimitRoundFinal = Reader.ReadInt8
                .PrizePoints = Reader.ReadInt16
                .LvlMin = Reader.ReadInt8
                .LvlMax = Reader.ReadInt8
                
                For b = 1 To NUMCLASES
                    .AllowedClasses(b) = Reader.ReadInt8
                Next b
                
                .InscriptionGld = Reader.ReadInt16
                .InscriptionGldPremium = Reader.ReadInt16
                
                .PrizeGld = Reader.ReadInt16
                .PrizeGldPremium = Reader.ReadInt16
                .PrizeObj.ObjIndex = Reader.ReadInt16
                .PrizeObj.Amount = Reader.ReadInt16
                
                .Config(eConfigEvent.eCascoEscudo) = Reader.ReadInt8
                
                .Config(eConfigEvent.eResu) = Reader.ReadInt8
                .Config(eConfigEvent.eInvisibilidad) = Reader.ReadInt8
                .Config(eConfigEvent.eOcultar) = Reader.ReadInt8
                .Config(eConfigEvent.eInvocar) = Reader.ReadInt8
            End If
        End With
    Next A

    
    If FrmPartidas.visible Then
        FrmPartidas.Tournaments_List
    End If
    
    #If ModoBig = 1 Then
        dockForm FrmPartidas.hWnd, FrmMain.PicMenu, True
    #Else
        Call FrmPartidas.Show(, FrmMain)
    #End If
    
End Sub

' # Estadísticas de un personaje
Private Sub HandleStatsUser()

    With InfoUser
        .UserName = Reader.ReadString8
        .Clase = Reader.ReadInt8
        .Raza = Reader.ReadInt8
        .Genero = Reader.ReadInt8
        .Elv = Reader.ReadInt8
        .Exp = Reader.ReadInt32
        .Elu = Reader.ReadInt32
        
        .Blocked = Reader.ReadInt8
        .BlockedHasta = Reader.ReadInt32
        
        .Gld = Reader.ReadInt32
        .Dsp = Reader.ReadInt32
        .Points = Reader.ReadInt32
        .Frags = Reader.ReadInt16
        .FragsCiu = Reader.ReadInt16
        .FragsCri = Reader.ReadInt16
        
        .Map = Reader.ReadInt16
        .X = Reader.ReadInt8
        .Y = Reader.ReadInt8
        
        .Hp = Reader.ReadInt16
    End With
    
    If MirandoStatsUser Then
        FrmStatsUser.Update_Info
    Else
        FrmStatsUser.Show , FrmMain
    End If
End Sub

' # Inventario de un personaje
Private Sub HandleStatsUser_Inventory()
    
    Dim IsBank As Boolean
    Dim A As Long
    
    With InfoUser.Inventory
        .NroItems = Reader.ReadInt8
        
        If .NroItems > 0 Then
            For A = 1 To .NroItems
                .Object(A).ObjIndex = Reader.ReadInt16
                .Object(A).Amount = Reader.ReadInt16
                .Object(A).Equipped = Reader.ReadInt8
            Next A
        
        End If
    End With
    
    ' # Actualizamos el inventario.
    FrmStatsUser.Load_Inventory (False)
End Sub
' # Banco de un personaje
Private Sub HandleStatsUser_Bank()
    
    Dim A As Long
    
    With InfoUser.Bank
        .NroItems = Reader.ReadInt8
        
        If .NroItems > 0 Then
            For A = 1 To .NroItems
                .Object(A).ObjIndex = Reader.ReadInt16
                .Object(A).Amount = Reader.ReadInt16
                .Object(A).Equipped = Reader.ReadInt8
            Next A
        
        End If
    End With
    
    ' # Actualizamos el inventario.
    FrmStatsUser.Load_Inventory (True)
End Sub
' # Hechizos de un personaje
Private Sub HandleStatsUser_Spells()
    
    Dim A As Long
    
    For A = 1 To MAXHECHI
        InfoUser.Spells(A) = Reader.ReadInt16
    Next A
    
    
    ' # Actualizamos la lista.
    FrmStatsUser.Load_Spells
    
End Sub

' # Skills de un personaje
Private Sub HandleStatsUser_Skills()
    
    Dim A As Long
    For A = 1 To NUMSKILLS
        InfoUser.Skills(A) = Reader.ReadInt16
    Next A
    
    ' # Actualizamos la lista.
    FrmStatsUser.Load_Skills
End Sub

' # Penas de un personaje
Private Sub HandleStatsUser_Penas()
    
    Dim A As Long
    
    InfoUser.PenasTime = Reader.ReadInt16
    InfoUser.PenasLast = Reader.ReadInt8
    
    If InfoUser.PenasLast > 0 Then
        ReDim InfoUser.Penas(1 To InfoUser.PenasLast) As String
        
        For A = 1 To InfoUser.PenasLast
            InfoUser.Penas(A) = Reader.ReadString8
        Next A
    End If
    
    ' # Actualizamos la lista.
    FrmStatsUser.Load_Penas
End Sub

' # Skins de un personaje
Private Sub HandleStatsUser_Skins()
    
    Dim A As Long
    
    InfoUser.Skins.Last = Reader.ReadInt8
    
    If InfoUser.Skins.Last > 0 Then
        ReDim InfoUser.Skins.ObjIndex(1 To InfoUser.Skins.Last) As Integer
        
        For A = 1 To InfoUser.Skins.Last
            InfoUser.Skins.ObjIndex(A) = Reader.ReadInt16
        Next A
    End If
    
    ' # Actualizamos la lista.
    FrmStatsUser.Load_Skins
End Sub


' # Bonificaciones de un personaje
Private Sub HandleStatsUser_Bonus()

    InfoUser.BonusLast = Reader.ReadInt8
    
    Dim A As Long
    
    If InfoUser.BonusLast > 0 Then
        ReDim InfoUser.BonusUser(1 To InfoUser.BonusLast) As UserBonus
            
        For A = 1 To InfoUser.BonusLast
            With InfoUser.BonusUser(A)
                .Tipo = Reader.ReadInt8
                .Value = Reader.ReadInt
                .Amount = Reader.ReadInt
                .DurationSeconds = Reader.ReadInt
                .DurationDate = Reader.ReadString8
            End With
        Next
    End If
    
    
    ' # Actualizamos la lista.
    FrmStatsUser.Load_Bonus
End Sub

' # Logros del personaje
Private Sub HandleStatsUser_Logros()

    
    FrmStatsUser.Load_Logros
End Sub


Private Sub HandleUpdateClient()

    MsgBox "El cliente se cerrará por actualización obligatoria. Entra nuevamente y lograrás acceder correctamente. Puede parecer tedioso, pero el re-ingreso es muy rápido y previene el uso de cheats en más de un 80%.", vbInformation
    prgRun = False
    'UpdateMd5File
End Sub

Public Sub WriteCastle(Optional ByVal CastleIndex As Byte = 0)
    Call Writer.WriteInt(ClientPacketID.Castle)
    Call Writer.WriteInt8(CastleIndex)
    Call modNetwork.Send(False)
End Sub

Public Sub WriteRequiredStatsUser(ByVal Tipo As Byte, ByVal Name As String)
    Call Writer.WriteInt(ClientPacketID.RequiredStatsUser)
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteString8(Name)
    Call modNetwork.Send(False)
End Sub
