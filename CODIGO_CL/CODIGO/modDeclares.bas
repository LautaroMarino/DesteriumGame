Attribute VB_Name = "Mod_Declaraciones"
'Exodo Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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
'Exodo Online is based on Baronsoft's VB6 Online RPG
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

Public Declare Function IsDebuggerPresent Lib "kernel32" () As Long

Public UsersOnline As Long
Public Const MAXUSERQUESTS As Integer = 30

Public Type tDropData

    ObjIndex As Integer
    Amount(1) As Integer
    Prob As Byte

End Type

Public Type tDrop

    Last As Byte
    data() As tDropData

End Type

Public DropLast                 As Integer

Public DropData()               As tDrop

Public keysMovementPressedQueue As clsArrayList

Public Const MAX_AURAS          As Byte = 5

Public Const BAR_BORDER         As Long = 84983

Public Const BAR_BACKGROUND     As Long = 84984

Public RenderizandoMap          As Boolean

Public RenderizandoIndex        As Long

Public CharIndex_MouseHover     As Long

Public NpcIndex_MouseHover      As Long

Public PuedeReclamar            As Boolean

Public Type tShopChars

    Name As String
    Dsp As Integer
    Gld As Long
    Elv As Byte
    Class As Byte
    Raze As Byte
    Head As Integer
    Hp As Integer
    Man As Integer
    Ups As Integer
    Porc As Byte
    
End Type

Public ShopCharLast     As Integer

Public ShopChars()      As tShopChars

Public MeditationUser() As Byte

Public MirandoShop      As Boolean

Public SelectedTienda   As Integer

Public EsModoEvento     As Byte

Public UsuariosOnline   As Integer

' // Formularios Escalados
    #If ModoBig = 1 Then
        
        #If FullScreen = 1 Then

            Public FrmMain As New frmMain_FullScreen
        #Else

            Public FrmMain As New frmMain_Scalled
        #End If

        Public FrmConnect_Account As New FrmConnect_AccountBig
    
    #Else
        
        #If FullScreen = 1 Then
            Public FrmMain            As New frmMain_FullScreen
        #Else
            Public FrmMain            As New frmMain_Classic
        #End If
        
        

        Public FrmConnect_Account As New FrmConnect_AccountOrig
    #End If


Public Enum eSubClass

    eMinero = 1
    ePescador = 2
    eTalador = 3

End Enum

Public SubClass        As eSubClass

Public Alias As String
Public CVU As String

Public LastDataAccount As String

Public LastDataPasswd  As String


    
Public ARENA_LAST      As Long
    
Public Type tBattleArena

    Name As String
    Maps() As Byte
    Limit As Byte
            
    ' Reseteable
    Users As Byte
            
End Type
    
Public Battle_Arenas() As tBattleArena


#If ModoBig = 1 Then

    Public Const SHAPE_LONGITUD       As Integer = 398

    Public Const SHAPE_LONGITUD_MITAD As Integer = 170

    Public Const WIDTH_EXP            As Integer = 194

#ElseIf ModoBig = 2 Then

    Public Const SHAPE_LONGITUD       As Integer = 198

    Public Const SHAPE_LONGITUD_MITAD As Integer = 60

    Public Const WIDTH_EXP            As Integer = 422

#Else

    Public Const SHAPE_LONGITUD       As Integer = 97

    Public Const WIDTH_EXP            As Integer = 122

    Public Const SHAPE_LONGITUD_MITAD As Integer = 97

#End If

Public Type t_packetCounters

    TS_CastSpell As Long
    TS_WorkLeftClick As Long
    TS_LeftClick As Long
    TS_UseItem As Long
    TS_UseItemU As Long
    TS_Walk As Long
    TS_Talk As Long
    TS_Attack As Long
    TS_Drop As Long
    TS_Work As Long
    TS_EquipItem As Long
    TS_GuildMessage As Long
    TS_QuestionGM As Long
    TS_ChangeHeading As Long

End Type

Public packetCounters             As t_packetCounters

Public Const CANT_PACKETS_CONTROL As Long = 400

Public Type t_packetControl

    last_count As Long
    'cant_iterations As Long = 10
    iterations(1 To 10) As Long

End Type

Public packetControl(1 To CANT_PACKETS_CONTROL) As t_packetControl

Public LastUpdateInv                            As Long

Public TempCharIndex                            As Integer

Public UserBankGold                             As Long

Public UserBankEldhir                           As Long
    
Public MapData9_Temp()                          As Byte

Public MapData10_Temp()                         As Byte

Public InitQuest                                As Boolean

Public SOURCE_INTERFACE                         As Long

Public Type RGBA

    b As Byte
    g As Byte
    r As Byte
    A As Byte

End Type

Public MiniMap_TileX        As Long

Public MiniMap_TileY        As Long

Public MiniMap_TileX_Actual As Long

Public MiniMap_TileY_Actual As Long


Public Const MAX_RETOS_CONFIG As Byte = 6
Public Const MAX_RETOS_TERRENO As Byte = 8

Public Enum eRetoConfig

    eInmovilizar = 0
    eResucitar = 1
    eEscudos = 2
    eCascos = 3
    eItems = 4
    eFuegoAmigo = 5

End Enum

Public Type tFight
    
    Tipo As Byte
    Users As String
    Config(0 To MAX_RETOS_CONFIG - 1) As eRetoConfig
    Rounds As Byte
    LimiteTiempo As Byte
    
    Gld As Long
    Dsp As Long
    Zona As Byte

End Type

Public ColourBody        As Long

Public SlotObjIndex      As Integer

Public MessagesSpam()    As String

Public MessagesSpam_Last As Integer

Public Const WM_USER = &H400

Public Const EM_GETSCROLLPOS = WM_USER + 221

Public Const EM_SETSCROLLPOS = WM_USER + 222

Public Const CP_UNICODE = 1200&

Public Const GT_USECRLF = 1&

Public Const GTL_USECRLF = 1&

Public Const GTL_PRECISE = 2&

Public Const GTL_NUMCHARS = 8&

Public Const EM_GETTEXTEX = WM_USER + 94

Public Const EM_GETTEXTLENGTHEX = WM_USER + 95

Public LastCapture         As String

Public CountDownTime       As Long

Public CountDownTime_Fight As Boolean

Public FX_SELECTED         As Integer

'Posicion en el Mundo
Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

'Posicion en un mapa
Public Type Position

    X As Long
    Y As Long

End Type

'Apunta a una estructura grhdata y mantiene la animacion
Public Type grh

    GrhIndex As Long
    Speed As Single
    started As Long
    Loops As Integer
    AnimacionContador As Single
    Alpha As Byte
    CantAnim As Long
    
    FxIndex As Integer

End Type

'Direcciones
Public Enum E_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As grh
    HeadOffset As Position
    BodyOffSet(E_Heading.NORTH To E_Heading.WEST) As Position

End Type 'Lista de cuerpos

Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As grh

End Type 'Lista de cabezas

Public Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As grh

End Type 'Lista de las animaciones de las armas

Public Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As grh

End Type 'Lista de las animaciones de los escudos

Public Type AuraData

    Walk(E_Heading.NORTH To E_Heading.WEST) As grh
    Color As Long

End Type 'Lista de las animaciones de las auras

Public GrhData()        As GrhData 'Guarda todos los grh
Public GrhDataDefault()        As GrhData ' Contiene todos los Grh sin alteraciones en calculos de x2

Public BodyData()       As BodyData

Public BodyDataAttack() As BodyData

Public HeadData()       As HeadData

Public FxData()         As tIndiceFx

Public WeaponAnimData() As WeaponAnimData

Public ShieldAnimData() As ShieldAnimData

Public CascoAnimData()  As HeadData

Public AuraAnimData()   As AuraData

Public Const MAX_FX     As Integer = 3

Public FX_LAST          As Integer

Public Enum eFX_Type

    FX_CONCENTRATION = 1 ' Slot dedicado para la meditación/concentraciones.
    FX_SPELL = 2 ' Slot Dedicado al hechizo.
    FX_EXTRA = 3 ' Slot utilizado para extras o un segundo hechizo.

End Enum

'Apariencia del personaje
Public Type Char

    Streamer As Boolean
    
    Craft As Integer
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    
    NowPosX As Long
    NowPosY As Long
    iBodyAttack As Integer
    Body As BodyData
    BodyAttack As BodyData
    Aura(1 To MAX_AURAS) As AuraData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    TimeAttackNpc As Long
    
    fX As grh
    FxIndex As Integer
    
    FxCount As Integer
    FxList() As grh
    
    Criminal As Byte
    Atacable As Boolean
    ColorNick As eNickColor
    NpcIndex As Integer
    Idle As Boolean
    Navegando As Boolean
    Speeding As Single
    
    GroupIndex As Byte
    
    ValidInvi As Boolean
    Nombre As String
    GuildName As String
    LastDialog As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    MovArmaEscudo As Boolean
    Moving As Boolean
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    OffsetY As Integer
    
    LastStep As Long
    Pie As Boolean
    Muerto As Boolean
    Invisible As Boolean
    Intermitencia As Boolean
    Priv As Byte
    
    ' RTree
    Width As Single
    Height As Single
    
    MinHp As Integer
    MaxHp As Integer
    MinMan As Integer
    MaxMan As Integer
    
    SoundSource As Long
    
    BarMin As Long
    BarMax As Long

End Type

'Info de un objeto
Public Type Obj

    ObjIndex As Integer
    Amount As Long
    SoundSource As Long ' Objects can make sounds too!
    
End Type

'Tipo de las celdas del mapa
Public Type MapBlock

    'light_value As RGBA
    ObjName As String
    
    Graphic(1 To 4) As grh
    CharIndex As Integer
    ObjGrh As grh
    
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    Damage As DList
    
    fX As grh
    FxIndex As Integer
    
    SoundSource As Long
    BarMin As Long
    BarMax As Long

End Type

'Info de cada mapa
Public Type MapInfo

    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer

End Type

Public CharList(1 To 10000) As Char

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()            As MapBlock ' Mapa

Public MapData_Copy()       As MapBlock ' Mapa

Public Enum E_BANK

    e_User = 1
    e_Account = 2

End Enum

Public SelectedBank              As E_BANK

'Api para generar un evento de tecla, en este caso Print Screen
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const FX_INMOVILIZAR      As Byte = 12

Public Const FX_APOCALIPSIS      As Byte = 13

Public Const FX_DESCARGA         As Byte = 11

Public Const FX_WARP             As Byte = 1

Public Const FX_TORMENTA         As Byte = 7

Public Const FX_REMOVER          As Byte = 0
                              
Public FightOn                   As Boolean

Public Fight_UserName            As String

Public LastDialogNpc             As Integer

Public Copy_HD                   As Long

Public Copy_MAC                  As String

Public CheckingDouble            As Long

Public CheckingDoubleValue(10)   As Long

Public CheckingDouble_U          As Long

Public CheckingDoubleValue_U(10) As Long

Public ControlActivated          As Boolean

Public HablaTemp                 As String

Public Enum eHabla

    e_Normal = 1
    e_Grito = 2
    e_Party = 3
    e_Clan = 4
    e_Susurro = 5

End Enum

Public IntervaloAnterior As Long

Public Enum eKeyPackets

    Key_UseItem = 0
    Key_UseSpell = 1
    Key_UseWeapon = 2

End Enum

Public Const MAX_KEY_PACKETS       As Byte = 2

Public KeyPackets(MAX_KEY_PACKETS) As Long

Public Type TempOnline

    UserName As String
    Value As Long

End Type

Public rNivel(1)              As TempOnline

Public rFrags(1)              As TempOnline

Public rRetos1(1)             As TempOnline

Public rRetos2(1)             As TempOnline

Public rTorneos(1)            As TempOnline

'Objetos públicos
Public DialogosClanes         As clsGuildDlg

Public Dialogos               As clsDialogs

Public Audio                  As clsAudio

Public InventorySkins As clsGrapchicalInventory

Public Inventario             As clsGrapchicalInventory

Public InvBanco(1)            As clsGrapchicalInventory

Public InvSkin                As clsGrapchicalInventory

Public InvMercader            As clsGrapchicalInventory

Public InvEvent               As clsGrapchicalInventory

' Inventarios del Mercado de Personajes
Public Mercader_Inv           As clsGrapchicalInventory

Public Mercader_Bank          As clsGrapchicalInventory

'Inventarios de comercio con usuario
Public InvComUsu              As clsGrapchicalInventory  ' Inventario del usuario visible en el comercio

Public InvOroComUsu(2)        As clsGrapchicalInventory  ' Inventarios de oro (ambos usuarios)

Public InvEldhirComUsu(2)     As clsGrapchicalInventory  ' Inventarios de oro (ambos usuarios)

Public InvOfferComUsu(1)      As clsGrapchicalInventory  ' Inventarios de ofertas (ambos usuarios)

Public InvComUser             As clsGrapchicalInventory

Public InvComNpc              As clsGrapchicalInventory  ' Inventario con los items que ofrece el npc

Public BlacksmithInv          As clsGrapchicalInventory

Public CustomKeys             As clsCustomKeys

Public CustomMessages         As clsCustomMessages

''
'The main timer of the game.
Public MainTimer              As clsTimer

'Error code
Public Const TOO_FAST         As Long = 24036

Public Const REFUSED          As Long = 24061

Public Const TIME_OUT         As Long = 24060

'Sonidos
Public Const SND_CLICK        As String = "click.Wav"

Public Const SND_PASOS1       As String = "23.Wav"

Public Const SND_PASOS2       As String = "24.Wav"

Public Const SND_NAVEGANDO    As String = "50.wav"

Public Const SND_OVER         As String = "click2.Wav"

Public Const SND_DICE         As String = "cupdice.Wav"

Public Const SND_LLUVIAINEND  As String = "lluviainend.wav"

Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

Public Enum eSound

    sConstruction = 237 ' Construcción de Crafting criatura
    sConquistCastle = 238 ' Sonido cuando está por conquistar el castillo.
    eDopaPerdida = 8 ' Sonido de la vaca

End Enum

' Head index of the casper. Used to know if a char is killed

' Variables de Intervalo

Public IntervaloUserPuedeAtacar      As Integer

Public IntervaloUserPuedeUsar        As Integer

Public IntervaloUserPuedeUsarClick   As Integer

Public IntervaloUpdatePos            As Integer

Public IntervaloUserPuedeCastear     As Integer

Public IntervaloUserPuedeShiftear    As Integer

Public IntervaloFlechasCazadores     As Integer

Public IntervaloCaminar              As Long

Public IntervaloHeading              As Long

Public Const CONST_INTERVALO_HEADING As Long = 120

Public IntervaloMagiaGolpe           As Integer

Public IntervaloGolpeMagia           As Integer

Public IntervaloGolpeUsar            As Integer

Public IntervaloUserPuedeTrabajar    As Integer

Public IntervalDrop                  As Integer

Public MacroBltIndex                 As Integer

Public Const CASPER_BODY_IDLE        As Integer = 829

Public Const TIME_CASPER_IDLE        As Long = 300


Public Const iCuerpoMuerto  As Integer = 8
Public Const iCuerpoMuerto_Legion  As Integer = 145
Public Const iCabezaMuerto  As Integer = 500
Public Const iCabezaMuerto_Legion  As Integer = 501

Public Const FRAGATA_FANTASMAL As Integer = 764
Public Const NUMATRIBUTES As Byte = 5

    Public Const HUMANO_H_PRIMER_CABEZA  As Integer = 502

    Public Const HUMANO_H_ULTIMA_CABEZA  As Integer = 546

    Public Const HUMANO_H_CUERPO_DESNUDO As Integer = 21

    Public Const ELFO_H_PRIMER_CABEZA    As Integer = 577

    Public Const ELFO_H_ULTIMA_CABEZA    As Integer = 608

    Public Const ELFO_H_CUERPO_DESNUDO   As Integer = 21

    Public Const DROW_H_PRIMER_CABEZA    As Integer = 639

    Public Const DROW_H_ULTIMA_CABEZA    As Integer = 669

    Public Const DROW_H_CUERPO_DESNUDO   As Integer = 32

    Public Const ENANO_H_PRIMER_CABEZA   As Integer = 700

    Public Const ENANO_H_ULTIMA_CABEZA   As Integer = 729

    Public Const ENANO_H_CUERPO_DESNUDO  As Integer = 53

    Public Const GNOMO_H_PRIMER_CABEZA   As Integer = 760

    Public Const GNOMO_H_ULTIMA_CABEZA   As Integer = 789

    Public Const GNOMO_H_CUERPO_DESNUDO  As Integer = 53

    '**************************************************
    Public Const HUMANO_M_PRIMER_CABEZA  As Integer = 547

    Public Const HUMANO_M_ULTIMA_CABEZA  As Integer = 576

    Public Const HUMANO_M_CUERPO_DESNUDO As Integer = 39

    Public Const ELFO_M_PRIMER_CABEZA    As Integer = 609

    Public Const ELFO_M_ULTIMA_CABEZA    As Integer = 638

    Public Const ELFO_M_CUERPO_DESNUDO   As Integer = 39

    Public Const DROW_M_PRIMER_CABEZA    As Integer = 670

    Public Const DROW_M_ULTIMA_CABEZA    As Integer = 699

    Public Const DROW_M_CUERPO_DESNUDO   As Integer = 40

    Public Const ENANO_M_PRIMER_CABEZA   As Integer = 730

    Public Const ENANO_M_ULTIMA_CABEZA   As Integer = 759

    Public Const ENANO_M_CUERPO_DESNUDO  As Integer = 60

    Public Const GNOMO_M_PRIMER_CABEZA   As Integer = 790

    Public Const GNOMO_M_ULTIMA_CABEZA   As Integer = 819

    Public Const GNOMO_M_CUERPO_DESNUDO  As Integer = 60

'Musica
Public Const MP3_Inicio As Byte = 101

Public ConsoleIndex     As Byte

Public RawServersList   As String

Public Type tColor

    r As Byte
    g As Byte
    b As Byte

End Type

Public ColoresPJ(0 To 50)                       As tColor

'CHOTS | Colores de diálogos customizables
Public Const MAXCOLORESDIALOGOS                 As Byte = 6

Public ColoresDialogos(1 To MAXCOLORESDIALOGOS) As tColor
'Referencias:
'1=Normal
'2=Clan
'3=Party
'4=Gritar
'5=Palabras Mágicas
'6=Susurrar
'CHOTS

Public CurServer                                As Integer

Public Enum ALINEACION_GUILD

    ALINEACION_LEGION = 1
    ALINEACION_NEUTRO = 2
    ALINEACION_ARMADA = 3
    ALINEACION_MASTER = 4

End Enum

Public CreandoClan       As Boolean

Public ClanName          As String

Public secClanName       As String

Public GuildAlineation   As ALINEACION_GUILD

Public Site              As String

Public UserCiego         As Boolean

Public UserEstupido      As Boolean

Public RainBufferIndex   As Long

Public FogataBufferIndex As Long

Public Const bCabeza = 1

Public Const bPiernaIzquierda = 2

Public Const bPiernaDerecha = 3

Public Const bBrazoDerecho = 4

Public Const bBrazoIzquierdo = 5

Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000

Public Const tUs = 600

Public Const PrimerBodyBarco = 84

Public Const UltimoBodyBarco = 87

Public NumEscudosAnims                                   As Integer

Public ItemsConstruibles()                               As tItemsConstruibles

Public Const MAX_BANCOINVENTORY_SLOTS                    As Byte = 40

Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

Public TradingUserName                                   As String

Public Const LoopAdEternum                               As Integer = 999

Public LastNpcSummoned                                   As Byte

'Objetos
Public Const MAX_INVENTORY_OBJS                          As Integer = 10000

''
' Cantidad de Skins por personaje
Public Const MAX_INVENTORY_SKINS                         As Byte = 50

Public Const MAX_INVENTORY_SLOTS                         As Byte = 30

Public Const MAX_NORMAL_INVENTORY_SLOTS                  As Byte = 30

Public Const MAX_NPC_INVENTORY_SLOTS                     As Byte = 30

Public Const MAXHECHI                                    As Byte = 35

Public Const INV_OFFER_SLOTS                             As Byte = 20

Public Const INV_GOLD_SLOTS                              As Byte = 1

Public Const MAXSKILLPOINTS                              As Byte = 100

Public Const MAXATRIBUTOS                                As Byte = 38

Public Const FLAGORO                                     As Integer = MAX_INVENTORY_SLOTS + 1

Public Const FLAGELDHIR                                  As Integer = FLAGORO + 1

Public Const GOLD_OFFER_SLOT                             As Integer = INV_OFFER_SLOTS + 1

Public Const ELDHIR_OFFER_SLOT                           As Integer = INV_OFFER_SLOTS + 2

Public Const FOgata                                      As Integer = 1521

Public Enum eFaccion

    fArmada = 1
    fLegion

End Enum

Public Enum eClass

    Mage = 1       'Mago
    Cleric = 2    'Clérigo
    Warrior = 3   'Guerrero
    Assasin = 4   'Asesino
    Bard = 5      'Bardo
    Druid = 6     'Druida
    Paladin = 7   'Paladín
    Hunter = 8   'Cazador
    Thief = 9     'Ladrón
    
    
    ' OFF
    Workerer = 10    'Trabajador

End Enum

Public Enum eCiudad

    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal

End Enum

Enum eRaza

    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano

End Enum

Public Enum eSkill

    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Apuñalar = 5
    Ocultarse = 6
    Talar = 7
    Defensa = 8
    Pesca = 9
    Mineria = 10
    Comercio = 11
    Domar = 12
    Proyectiles = 13
    Navegacion = 14
    Resistencia = 15

End Enum

Public Enum eAtributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5

End Enum

Enum eGenero

    Hombre = 1
    Mujer

End Enum

Public Enum PlayerType

    User = &H1
    SemiDios = &H2
    Libre = &H4
    Dios = &H8
    Admin = &H10
    ChaosCouncil = &H20
    RoyalCouncil = &H40

End Enum

Public Enum eOBJType

    otUseOnce = 1
    otWeapon = 2
    otarmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otGemas = 35
    otArbolElfico = 36
    otMochilas = 37
    otYacimientoPez = 38
    otGuitaDsp = 39
    otRangeQuest = 40
    otTeleportInvoker = 41
    otPendienteParty = 42
    otItemNpc = 43      ' Drops de los npcs
    otItemRandom = 52
    otCofreAbierto = 53
    otcofre = 54
    otGemaTelep = 55
    otGemasEffect = 56
    otMonturas = 57
    otReliquias = 58
    otAuras = 59
    oteffect = 60 ' Objetos que dan efectos sobre el personaje (Exp, Oro, etc)
    otMagic = 61 ' Objetos tales como Anillo Mágico y Laud Mágico
    otTransformVIP = 62
    otTravel = 63
    otLibroGuild = 64
    otActaNick = 65 ' Acta de Nacimiento:: Cambia de Nombre
    otActaGuild = 66 ' Escrituras del Clan :: Cambia el Lider por Otro (Accion para el Lider)
    otCualquiera = 1000

End Enum

Public Enum eMochilas

    Mediana = 1
    GRANDE = 2

End Enum

Public MaxInventorySlots As Byte

Public Const FundirMetal As Integer = 88

Public Const TeleportInvoker = 99

' Determina el color del nick
Public Enum eNickColor

    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
    ieCastleGuild = &H8
    ieCastleUser = &H10
    ieCAOS = &H20
    ieArmada = &H40
    ieShield = &H80

End Enum

Public Enum eGMCommands

    GMMessage = 1           '/GMSG
    ShowName                '/SHOWNAME
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    GoToChar                '/IRA
    Invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    RequestCharInfo         '/INFO
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IpToNick                '/IP2NICK
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    DoBackUp                '/DOBACKUP
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
    ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
    ChangeMapInfoLvl        '/MODMAPINFO NIVELMIN
    ChangeMapInfoLimpieza   '/MODMAPINFO LIMPIEZA
    ChangeMapInfoItems      '/MODMAPINFO ITEMS
    ChangeMapInfoExp         '/MODMAPINFO EXP
    ChangeMapInfoattack     '/MODMAPINFO ATAQUE
    SaveChars               '/GRABAR
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    EnableDenounces         '/DENUNCIAS
    ShowDenouncesList       '/SHOW DENUNCIAS
    MapMessage              '/MAPMSG
    SetDialog               '/SETDIALOG
    Impersonate             '/IMPERSONAR
    Imitate                 '/MIMETIZAR
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest

    SearchObj
    SolicitaSeguridad
    CheckingGlobal
    CountDown
    GiveBackUser
    
    Pro_Seguimiento         ' Dejamos de seguir al personaje elegido
    
    Events_KickUser
    
    SendDataUser
    SearchDataUser
    ChangeModoArgentum
    StreamerBotSetting
    LotteryNew
End Enum
        
'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_DRAG_ON                           As String = "SEGURO DE DRAG ACTIVADO"

Public Const MENSAJE_DRAG_OFF                          As String = "SEGURO DE DRAG DESACTIVADO"

Public Const MENSAJE_CRIATURA_FALLA_GOLPE              As String = "¡¡¡La criatura falló el golpe!!!"

Public Const MENSAJE_CRIATURA_MATADO                   As String = "¡¡¡La criatura te ha matado!!!"

Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO             As String = "¡¡¡Has rechazado el ataque con el escudo!!!"

Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO     As String = "¡¡¡El usuario rechazó el ataque con su escudo!!!"

Public Const MENSAJE_FALLADO_GOLPE                     As String = "¡¡¡Has fallado el golpe!!!"

Public Const MENSAJE_SEGURO_ACTIVADO                   As String = ">>SEGURO ACTIVADO<<"

Public Const MENSAJE_SEGURO_DESACTIVADO                As String = ">>SEGURO DESACTIVADO<<"

Public Const MENSAJE_SEGURO_RETOS_ACTIVADO             As String = ">>PANEL DE INVITACIÓN DE RETOS ACTIVADO<<"

Public Const MENSAJE_SEGURO_RETOS_DESACTIVADO          As String = ">>PANEL DE INVITACIÓN DE RETOS DESACTIVADO<<"

Public Const MENSAJE_PIERDE_NOBLEZA                    As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."

Public Const MENSAJE_USAR_MEDITANDO                    As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON                    As String = ">>SEGURO DE RESURRECCION ACTIVADO<<"

Public Const MENSAJE_SEGURO_RESU_OFF                   As String = ">>SEGURO DE RESURRECCION DESACTIVADO<<"

Public Const MENSAJE_GOLPE_CABEZA                      As String = "¡¡La criatura te ha pegado en la cabeza por "

Public Const MENSAJE_GOLPE_BRAZO_IZQ                   As String = "¡¡La criatura te ha pegado el brazo izquierdo por "

Public Const MENSAJE_GOLPE_BRAZO_DER                   As String = "¡¡La criatura te ha pegado el brazo derecho por "

Public Const MENSAJE_GOLPE_PIERNA_IZQ                  As String = "¡¡La criatura te ha pegado la pierna izquierda por "

Public Const MENSAJE_GOLPE_PIERNA_DER                  As String = "¡¡La criatura te ha pegado la pierna derecha por "

Public Const MENSAJE_GOLPE_TORSO                       As String = "¡¡La criatura te ha pegado en el torso por "

Public Const MENSAJE_DRAG_DESACTIVADO                  As String = ">>SEGURO DRAG & DROP DESACTIVADO<<"

Public Const MENSAJE_DRAG_ACTIVADO                     As String = ">>SEGURO DRAG & DROP ACTIVADO<<"

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1                                 As String = "¡¡"

Public Const MENSAJE_2                                 As String = "!!"

Public Const MENSAJE_11                                As String = "¡"

Public Const MENSAJE_22                                As String = "!"

Public Const MENSAJE_GOLPE_CRIATURA_1                  As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO                      As String = " te atacó y falló!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA             As String = " te ha pegado en la cabeza por "

Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ          As String = " te ha pegado el brazo izquierdo por "

Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER          As String = " te ha pegado el brazo derecho por "

Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ         As String = " te ha pegado la pierna izquierda por "

Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER         As String = " te ha pegado la pierna derecha por "

Public Const MENSAJE_RECIVE_IMPACTO_TORSO              As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1                 As String = "¡¡Le has pegado a "

Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA            As String = " en la cabeza por "

Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ         As String = " en el brazo izquierdo por "

Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER         As String = " en el brazo derecho por "

Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ        As String = " en la pierna izquierda por "

Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER        As String = " en la pierna derecha por "

Public Const MENSAJE_PRODUCE_IMPACTO_TORSO             As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA                     As String = "Haz click sobre el objetivo..."

Public Const MENSAJE_TRABAJO_PESCA                     As String = "Haz click sobre el sitio donde quieres pescar..."

Public Const MENSAJE_TRABAJO_ROBAR                     As String = "Haz click sobre la víctima..."

Public Const MENSAJE_TRABAJO_TALAR                     As String = "Haz click sobre el árbol..."

Public Const MENSAJE_TRABAJO_MINERIA                   As String = "Haz click sobre el yacimiento..."

Public Const MENSAJE_TRABAJO_FUNDIRMETAL               As String = "Haz click sobre la fragua..."

Public Const MENSAJE_TELEPORT_INVOKER                  As String = "Haz click sobre una posición válida..."

Public Const MENSAJE_TRABAJO_PROYECTILES               As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1                    As String = "Si deseas entrar en una party con "

Public Const MENSAJE_ENTRAR_PARTY_2                    As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE                              As String = "Cantidad de NPCs: "

Public Const MENSAJE_FRAGSHOOTER_TE_HA_MATADO          As String = "te ha matado!"

Public Const MENSAJE_FRAGSHOOTER_HAS_MATADO            As String = "Has matado a"

Public Const MENSAJE_FRAGSHOOTER_HAS_GANADO            As String = "Has ganado "

Public Const MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA As String = "puntos de experiencia."

Public Const MENSAJE_NO_VES_NADA_INTERESANTE           As String = "No ves nada interesante."

Public Const MENSAJE_HAS_MATADO_A                      As String = "Has matado a "

Public Const MENSAJE_HAS_GANADO_EXPE_1                 As String = "Has ganado "

Public Const MENSAJE_HAS_GANADO_EXPE_2                 As String = " puntos de experiencia."

Public Const MENSAJE_TE_HA_MATADO                      As String = " te ha matado!"

Public Const MENSAJE_HOGAR                             As String = "Has llegado a tu hogar. El viaje ha finalizado."

Public Const MENSAJE_HOGAR_CANCEL                      As String = "Tu viaje ha sido cancelado."

Public Const MENSAJE_MODOSTREAM_ON1                    As String = "Has activado el MODO STREAM. ¿No estás strimeando? Desactivalo con /STREAM"

Public Const MENSAJE_MODOSTREAM_ON2                    As String = "Utiliza el comando /STREAMLINK y con un espacio de por medio ingresa el LINK de tu canal. Recuerda que cualquier link inválido podría ser tomado como pena sobre tu cuenta y/o personajes."

Public Const MENSAJE_MODOSTREAM_ON3                    As String = "¡Disfrutalo! Y esperamos desde nuestra comunidad que puedas sumar muchos seguidores."

Public Const MENSAJE_MODOSTREAM_OFF                    As String = "MODO STREAM desactivado. ¡Te esperamos la próxima!"

Public Enum eMessages

    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    ResuscitationSafeOff
    ResuscitationSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    HaveKilledUser
    UserKill
    EarnExp
    GoHome
    CancelGoHome
    FinishHome
    DragSafeOn
    DragSafeOff
    ModoStreamOn
    ModoStreamOff

End Enum

'Inventario
Type Inventory

    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    ObjType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxDefMag As Integer
    MinDefMag As Integer
    MaxHitMag As Integer
    MinHitMag As Integer
    
    MaxHit As Integer
    MinHit As Integer
    ValorAzul As Long
    CanUse As Boolean

    Bronce As Byte
    Plata As Byte
    Oro As Byte
    Premium As Byte
    
    ExistSkin As Integer ' # Identifica si tiene la skin
End Type

Type NpCinV

    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Integer
    Valor As Single
    ValorAzul As Single
    CanUse As Boolean
    ObjType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    MaxDefMag As Integer
    MinDefMag As Integer
    MaxHitMag As Integer
    MinHitMag As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String

    Animation As Integer

End Type

Type tItemsConstruibles_Required

    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    Amount As Long

End Type

Enum eItemsConstruibles_Subtipo

    eArmadura = 1
    eCasco = 2
    eEscudo = 3
    eArmas = 4
    eMuniciones = 5
    eEmbarcaciones = 6
    eObjetoMagico = 7
    eInstrumento = 8
    
    eFundicion = 9 ' Reservado para fundicion

End Enum

Public Nombres As Boolean

Public Enum TypeWorking

    eValidation = 1
    eRecover = 2
    eKill = 3

End Enum

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS)   As Inventory

Public UserHechizos(1 To MAXHECHI)                As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV

Public UserMeditar                                As Boolean

Public UserName                                   As String

Public PasswdGM                                   As String

Public UserValidation                             As Long

Public TipeWorking                                As TypeWorking

Public UserIpExternal                             As String

Public UserPin                                    As String

Public UserKey                                    As String

Public UserEmail                                  As String

Public GroupIndex                                 As Byte

Public uName                                      As String 'nick con caracteres especiales

Public UserPassword                               As String

Public UserMaxMAN                                 As Long

Public UserMinMAN                                 As Long

Public UserMaxSTA                                 As Integer

Public UserMinSTA                                 As Integer

Public UserMaxAGU                                 As Byte

Public UserMinAGU                                 As Byte

Public UserMaxHAM                                 As Byte

Public UserMinHAM                                 As Byte

Public UserGLD                                    As Long

Public UserDSP                                    As Long

Public UserPoints                                 As Long

Public UserLvl                                    As Integer

Public UserPort                                   As Integer

Public UserServerIP                               As String

Public UserEstado                                 As Byte '0 = Vivo & 1 = Muerto

Public UserPasarNivel                             As Long

Public UserExp                                    As Long

Public UserEstadisticas                           As tEstadisticasUsu

Public UserDescansar                              As Boolean

Public Moviendose                                 As Boolean

Public UserMaxHP                                  As Long

Public UserMinHP                                  As Long

Public pausa                                      As Boolean

Public UserParalizado                             As Boolean

Public UserEnvenenado                             As Boolean

Public UserEvento                                 As Boolean

Public UserNavegando                              As Boolean

Public UserMontando                               As Boolean

Public UserFuerza                                 As Byte

Public UserAgilidad                               As Byte

Public UserWeaponEqpSlot                          As Byte

Public UserAnilloEqpSlot                          As Byte

Public UserMagicEqpSlot                           As Byte

Public UserArmourEqpSlot                          As Byte

Public UserHelmEqpSlot                            As Byte

Public UserShieldEqpSlot                          As Byte

Public UserLeader                                 As Boolean

Public UserMuerto                                 As Boolean

'<-------------------------NUEVO-------------------------->
Public Comerciando                                As Boolean


Public MirandoRetos                               As Boolean
Public MirandoObjetos                             As Boolean

Public MirandoPartidas As Boolean

Public MirandoOpcionesNpc                         As Boolean

Public MirandoForo                                As Boolean

Public MirandoParty                               As Boolean

Public MirandoEstadisticas                        As Boolean

Public MirandoCantidad                            As Boolean
        
Public MirandoRank                                As Boolean

Public MirandoGuildPanel                          As Boolean

Public MirandoTravel                              As Boolean

Public MirandoComerciarUsu                        As Boolean

Public MirandoListaDrops                          As Boolean

Public MirandoListaCofres                         As Boolean

Public MirandoSkins                               As Boolean

Public MirandoBanco                               As Boolean

Public MirandoConcentracion                       As Boolean

Public MirandoComerciar                           As Boolean

Public MirandoCuenta                              As Boolean


Public MirandoObj                                 As Boolean
Public MirandoNpc                                 As Boolean
Public MirandoMercader                            As Boolean
Public MirandoOffer                               As Boolean
Public MirandoStatsUser                           As Boolean
Public EsperandoValidacion                        As Boolean 'MAO

'<-------------------------NUEVO-------------------------->

Public UserClase                                  As eClass

Public UserSexo                                   As eGenero

Public UserFaccion                                As eFaccion

Public WorkRaza                                   As eRaza

Public WorkGenero                                 As eGenero

Public WorkName                                   As String

Public UserRaza                                   As eRaza

Public Const NUMCIUDADES                          As Byte = 5

Public Const NUMSKILLS                            As Byte = 15

Public Const NUMSKILLSESPECIAL                    As Byte = 9

Public Const NUMATRIBUTOS                         As Byte = 5


Public Const NUMCLASES As Byte = 9


Public Const NUMRAZAS                    As Byte = 5

Public UserSkills(1 To NUMSKILLS)        As Byte

Public SkillsNames(1 To NUMSKILLS)       As String

Public UserAtributos(1 To NUMATRIBUTOS)  As Byte

Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES)        As String

Public ListaRazas(1 To NUMRAZAS)         As String

Public ListaRazasShort(1 To NUMRAZAS)    As String

Public ListaClases(1 To NUMCLASES)       As String

Public SkillPoints                       As Integer

Public Alocados                          As Integer

Public Flags()                           As Integer

Public Oscuridad                         As Integer

Public logged                            As Boolean

Public UsingSkill                        As Integer

Public MD5HushYo                         As String * 16

Public EsPartyLeader                     As Boolean

Public ServerSelected                    As Byte

Public Enum E_MODO

    e_LoginAccount = 1
    e_LoginAccountNew = 2
    e_LoginAccountNewChar = 3
    e_LoginAccountChar = 4
    e_LoginAccountRemove = 5
    e_LoginAccountRecover = 6
    e_LoginAccountPasswd = 7
    e_LoginMercaderOff = 8
    e_LoginName = 9
    
    e_DisconnectForced = 17

End Enum

Public EstadoLogin As E_MODO
   
Type tItemsConstruibles

    Loaded As Boolean
    
    Name As String
    GrhIndex As Long
    RequiredCant As Byte
    ObjIndex As Integer
    Amount As Long
    
    ObjType As eOBJType
    SubType As eItemsConstruibles_Subtipo
    CanUse As Boolean
    RequiredPremium As Long
    
    Required() As tItemsConstruibles_Required
    
    Bronce As Byte
    Plata As Byte
    Oro As Byte
    Premium As Byte
    LvlMin As Byte
    LvlMax As Byte
    
    MinDef As Integer
    MaxDef As Integer
    Ropaje As Integer
    
    MinHit As Integer
    MaxHit As Integer
    StaffDamageBonus As Integer
    
    MinDefMag As Integer
    MaxDefMag As Integer
    
    Weapon As Integer
    Shield As Integer
    Helm As Integer
    
    NoSeCae As Byte
    Envenena As Byte
    ClassValid(1 To NUMCLASES) As Byte
    ValidTotal As Boolean
    
    EdicionLimitada As Byte
    Text As String
    Heading As E_Heading

End Type

Public Enum FXIDs

    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARXXGRANDE = 16
    FXMEDITARXXXGRANDE = 70
    
    FXSANGRE = 14
    FXSWING = 14
    
    FX_INCINERADO = 72
    
    FX_LEVEL = 1 ' Pasaje de Nivel
    FX_APUÑALADA = 14 ' Efecto al Apuñalar

End Enum

Public Enum eClanType

    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal

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

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
' @param ZONAOSCURA lo que haya en este trigger no será visible
' @param CASA todo lo que tenga este trigger forma parte de una casa
'
Public Enum eTrigger

    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    ZONAOSCURA = 7
    CASA = 8
    AutoResu = 9
    LavaActiva = 10

End Enum

'Server stuff
Public RequestPosTimer   As Integer 'Used in main loop

Public stxtbuffer        As String 'Holds temp raw data from server

Public stxtbuffercmsg    As String 'Holds temp raw data from server

Public SendNewChar       As Boolean 'Used during login

Public Connected         As Boolean 'True when connected to server

Public DownloadingMap    As Boolean 'Currently downloading a map from server

Public UserMap           As Integer

Public UserMapName       As String

Public Map_TimeRender    As Integer

'Control
Public prgRun            As Boolean 'When true the program ends

Public IPdelServidor     As String

Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

'para escribir y leer variables
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpFileName As String) As Long

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
    
    BodyOffSetX(1 To 4) As Integer
    BodyOffSetY(1 To 4) As Integer

End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer

End Type

Public EsperandoLevel            As Boolean

'FragShooter variables
Public FragShooterCapturePending As Boolean

Public FragShooterNickname       As String

Public FragShooterKilledSomeone  As Boolean

Public Traveling                 As Boolean

Public Const OFFSET_HEAD         As Integer = -34

Public Enum eSMType

    sResucitation
    sSafemode
    mSpells
    mWork
    DragMode

End Enum

Public Declare Function MapVirtualKey _
               Lib "user32" _
               Alias "MapVirtualKeyA" (ByVal wCode As Long, _
                                       ByVal wMapType As Long) As Long

Public Declare Sub keybd_event _
               Lib "user32" (ByVal bVk As Byte, _
                             ByVal bScan As Byte, _
                             ByVal dwFlags As Long, _
                             ByVal dwExtraInfo As Long)

Public Const VK_MENU = &H12

Public Const VK_SNAPSHOT = &H2C

Public Const KEYEVENTF_KEYUP = &H2

Public IsSeguro                As Boolean

Public Const SM_CANT           As Byte = 4

Public SMStatus(SM_CANT)       As Boolean

Public ClientMeditation        As Byte

Public Const GRH_MEDITACION    As Long = 84055

'Hardcoded grhs and items
Public Const GRH_INI_SM        As Integer = 4978

Public Const ORO_INDEX         As Integer = 12

Public Const ELDHIR_INDEX      As Integer = 943

Public Const ORO_GRH           As Integer = 511

Public Const ELDHIR_GRH        As Integer = 16361

Public Const GRH_HALF_STAR     As Integer = 5357

Public Const GRH_FULL_STAR     As Integer = 5358

Public Const GRH_GLOW_STAR     As Integer = 5359

Public Const LH_GRH            As Integer = 724

Public Const LP_GRH            As Integer = 725

Public Const LO_GRH            As Integer = 723

Public Const MADERA_GRH        As Integer = 550

Public Const MADERA_ELFICA_GRH As Integer = 4803

Public picMouseIcon            As Picture

Public Enum eMoveType

    Inventory = 1
    Bank

End Enum

Public Const MP3_INITIAL_INDEX As Integer = 1000

'#######
' MENUES
Public Type tMenuAction

    NormalGrh As Integer
    FocusGrh As Integer
    ActionIndex As Byte

End Type

Public Type tMenu

    NumActions As Byte
    Actions() As tMenuAction

End Type

Public MenuInfo() As tMenu

Public Enum eMenuAction

    ieCommerce = 1
    iePriestHeal
    ieHogar
    iePetStand
    iePetFollow
    ieReleasePet
    ieTrain
    ieSummonLastNpc
    ieRestToggle
    ieBank
    ieFactionEnlist
    ieFactionReward
    ieFactionWithdraw
    ieFactionInfo
    ieGamble
    ieBlacksmith
    ieMakeLingot
    ieMeltDown
    ieShareNpc
    ieStopSharingNpc
    ieTameNpc
    ieMakeFireWood
    ieLightFire

End Enum

'Caracteres
Public Const CAR_ESPECIALES = "áàäâÁÀÄÂéèëêÉÈËÊíìïîÍÌÏÎóòöôÓÒÖÔúùüûÚÙÜÛñÑ'"

Public Const CAR_COMUNES = "aaaaAAAAeeeeEEEEiiiiIIIIooooOOOOuuuuUUUUnN "

Public Const CAR_ESPECIALES_CLANES = "(.;,"

Public Const CAR_COMUNES_CLANES = "$$$$"

' GRUPOS
Public Const MAX_MEMBERS_GROUP         As Byte = 5

Public Const MAX_REQUESTS_GROUP        As Byte = 10

Private Const MAX_GROUPS               As Byte = 100

Private Const SLOT_LEADER              As Byte = 1

Private Const EXP_BONUS_MAX_MEMBERS    As Single = 1.05 '%

Private Const EXP_BONUS_LEADER_PREMIUM As Single = 1.05 '%

Public Type tUserGroup

    Name As String
    Exp As Long
    PorcExp As Byte
    
End Type

Public Type tGroups

    Members As Byte
    User(1 To MAX_MEMBERS_GROUP) As tUserGroup
    Requests(1 To MAX_REQUESTS_GROUP) As String

End Type

Public Groups As tGroups

Public Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Public Declare Function RedrawWindow _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal lprcUpdate As Long, _
                             ByVal hrgnUpdate As Long, _
                             ByVal fuRedraw As Long) As Long

Public Declare Function SetPixel _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long) As Long

Public Enum PlayLoop

    plNone = 0

End Enum

Public Type tCharCommerce

    Body As Integer
    Head As Integer
    Helm As Integer
    Shield As Integer
    Weapon As Integer
    
    LastSlot As Byte
    
    Slot(MAX_NPC_INVENTORY_SLOTS) As Byte

End Type

Public CharCommerce              As tCharCommerce

' Clanes

Public Selected_GuildIndex As Integer

Public Const MAX_GUILDS          As Integer = 300

Public Const MAX_GUILD_MEMBER    As Byte = 30

Public Const MAX_GUILD_CODEX     As Byte = 4

Public Const MIN_GUILD_POINTS    As Integer = 2000

Public Const MAX_GUILD_LEN       As Byte = 15

Public Const MAX_GUILD_LEN_CODEX As Byte = 60

Public Const MAX_GUILD_RANGE     As Byte = 4

Public Enum eGuildRange

    rNone = 0
    rFound = 1
    rLeader = 2
    rVocero = 3

End Enum

Public Enum eGuildAlineation

    a_Neutral = 0
    a_Armada = 1
    a_Legion = 2

End Enum

Public GuildLevel                             As Integer

Public MiniMap_Friends(1 To MAX_GUILD_MEMBER) As WorldPos

' Info del clan seleccionado
Public Type tGuildMemberInfo

    Name As String
    Range As eGuildRange
    Elv As Byte
    Class As Byte
    Raze As Byte
    Body As Integer
    Head As Integer
    
    Helm As Integer
    Shield As Integer
    Weapon As Integer
    
    Points As Long

End Type

Public Type tGuild

    Index As Integer
    
    Name As String
    Alineation As eGuildAlineation

    Member As Byte
    MaxMember As Byte
    Members(1 To MAX_GUILD_MEMBER) As tGuildMemberInfo

    Colour As Long
    Lvl As Byte
    Exp As Long
    Elu As Long

End Type

' Solapas de la GUI
Public Enum rGuildPanel

    rList = 0       ' Lista de Clanes
    rEvents = 1     ' Lista de Eventos
    rFound = 2      ' Fundar Nuevo Clan
    rLeader = 3     ' Panel del Lider
    rPerfil = 4     ' Perfil de un Clan

End Enum



Public Const GUILD_CRISTAL    As Integer = 2206


Public TempAlineation                   As eGuildAlineation

Public CastleGuild(2)                   As Integer    ' Clanes de Castillos

Public GuildPanel                       As rGuildPanel 'Solapa Seleccionada

Public GuildSelected                    As Integer ' Clan seleccionado para visualizar la información

Public LastSelectedGuild                As Integer ' Scroll de los Clanes

Public GuildsInfo(1 To MAX_GUILDS)      As tGuild

Public Type tDrops

    Name As String
    ObjIndex As Integer
    Amount As Long
    Probability As Byte

End Type

Public Const MAX_NPC_DROPS As Byte = 5

Public Type tNpc

    Name As String
    Body As Integer
    Head As Integer
    cant As Integer
    Exp As Long
    Gld As Long
    Eldhir As Long
    
    NroSpells As Byte
    Spells() As String
    Drop(1 To MAX_NPC_DROPS) As tDrops
    Hp As Long
    
    MinHit As Integer
    MaxHit As Integer

End Type

Public Type tViajes
    
    Name As String
    Maps() As Integer
    Duration As Long
    Gld As Long
    Eldhir As Long
    
    ' Teletransportación
    Map As Integer
    X As Integer
    Y As Integer
    
    Npcs() As tNpc
    SelectedNpc As Byte
    
    LvlMin As Byte
    LvlMax As Byte

End Type

Public TravelSelected                     As Integer

Public TravelNpc()                        As tViajes

Public LastCharIndex                      As Long

' CUENTAS

Public CreandoPersonaje                   As Boolean

Public CreandoPersonaje_Slot              As Byte

Public Const ACCOUNT_MAX_CHARS            As String = 10

Public Const ACCOUNT_MIN_CHARACTER_CHAR   As Byte = 3

Public Const ACCOUNT_MAX_CHARACTER_CHAR   As Byte = 15

Public Const ACCOUNT_MIN_CHARACTER_KEY    As Byte = 20

Public Const ACCOUNT_MIN_CHARACTER_PASSWD As Byte = 8

Public Const MERCADER_MAX_LIST            As Integer = 255 '

Public Const MERCADER_MAX_GLD             As Long = 2000000000 ' 2.000.000.000

Public Const MERCADER_MAX_DSP             As Long = 100000 '100.000

Public Const MERCADER_GLD_SALE            As Long = 1500 ' 1.000 pide de base de Monedas de oro.

Public Const MERCADER_MIN_LVL             As Byte = 15    ' Pide Nivel 15 para poder ser publicado.

Public Const MERCADER_MAX_OFFER           As Byte = 50

Public Const MERCADER_MAX_SALE            As Byte = 50

Public Type UserOBJ

    Name As String
    GrhIndex As Long
    
    ObjIndex As Integer
    Amount As Long
    Equipped As Byte
    CanUse As Boolean
End Type

Public Type tMercaderChar

    Name As String
    Guild As String
    
    Body As Integer
    Head As Integer
    Weapon As Integer
    Shield As Integer
    Helm As Integer
    
    Elv As Byte
    Exp As Long
    Elu As Long
    
    Hp As Integer
    Constitucion As Integer
    
    Class As Byte
    Raze As Byte
    
    Desc As String  ' Muestra la información
    DescShort As String
    
    Faction As Byte
    FactionRange As Byte
    FragsCiu As Integer
    FragsCri As Integer
    Gld As Long
    
    Object(1 To MAX_INVENTORY_SLOTS) As Obj
    Bank(1 To MAX_BANCOINVENTORY_SLOTS) As Obj
    Spells(1 To 35) As String
    Skills(1 To NUMSKILLS) As Byte

End Type

Public Type tMercader

    Loaded As Boolean
    ID As Integer
    Account As String
    Blocked As Byte
    Gld As Long
    Dsp As Long
    Desc As String
    
    Char As Byte
    Chars(1 To ACCOUNT_MAX_CHARS) As tMercaderChar
    bChars(1 To ACCOUNT_MAX_CHARS) As Byte
    IDCHARS(1 To ACCOUNT_MAX_CHARS) As Byte
    Slot As Integer

End Type

Public MercaderOff                               As Byte ' Solicita quitar o listar ofertas, estando off pj pero on cuenta, por eso acá.
Public MercaderGld                               As Long      ' Oro Temporal de cuanto te va a salir la publicacion!!!

Public MercaderUser                              As tMercader
Public MercaderUserOffer                          As tMercader

Public MercaderChars(1 To ACCOUNT_MAX_CHARS)     As tMercaderChar

Public MercaderList(0 To MERCADER_MAX_LIST)      As tMercader
Public MercaderListOffer(0 To MERCADER_MAX_LIST) As tMercader

Public MercaderList_Copy(0 To MERCADER_MAX_LIST) As tMercader

Public MercaderSelected As Integer  ' Publi seleccionado
Public MercaderSelected1 As Integer ' Publicación seleccionada una vez que confirma la oferta
Public MercaderSelectedOffer1 As Integer ' Oferta seleccionada una vez que el vendedor tiene los botones habilitados de  ACEPTAR/RECHAZAR

Public MercaderSelectedOffer As Integer  ' Oferta seleccionada


Public MercaderUserSlot As Integer ' Slot de MERCADO que tiene el usuario para su cuenta. En caso de tener publicacion vigente.

Public Enum eMercaderSelected

    ePanelInitial = 0 ' Panel Inicial. Dibuja todos los Nicks de los Comerciantes & Npcs Interactivos
    ePanelMercaderList = 1 ' Muestra la Lista de de Ultimas ventas que se publicaron.
    ePanelSearch = 2 ' Buscador de Publicaciones a través de Nombre.
    ePanelAuction = 3 ' Panel de las Subastas
    ePanelApuestas = 4 ' Panel de Apuestas (Torneos Automáticos)
    ePanelPublication = 5 ' Panel de Publicación
    ePanelValidation = 6 ' Esperando un código de validacion...
    ePanelOffer = 6 ' Panel de Ofertas

End Enum

'Public MercaderSelected As eMercaderSelected

Public Enum eMercaderPJSelected

    ePanelInv = 1
    ePanelBank = 2
    ePanelSpell = 3
    ePanelSkills = 4

End Enum

Public Mercader_ModoOferta As Boolean

Public MercaderPJ_Selected As eMercaderPJSelected

Public MercaderID          As Integer    ' ID de la selección de Slot

Public MercaderPJ          As Integer    ' ID para scrollear la lista de pj

Public MercaderID_Selected As Integer 'ID para scrollear la lista

Public MercaderLoaded      As Boolean ' Determina que el usuario está solicitando más info de la que tiene..

Public Type tAccountChar

    ID As Integer
    Name As String
    Guild As String
    
    Body As Integer
    Head As Integer
    Helm As Integer
    Shield As Integer
    Weapon As Integer
    
    Elv As Byte
    Exp As Long
    Elu As Long
    
    Class As Byte
    Raze As Byte
    Gld As Long
    GldBank As Long
    Eldhir As Long
    
    Ban As Byte
    NumPenas As Byte
    Penas() As String
    
    StatusBronce As Byte
    StatusPlata As Byte
    StatusOro As Byte
    StatusPremium As Byte
    
    Faction As Byte
    FactionRange As Byte
    FragsCiu As Integer
    FragsCri As Integer
    FragsOther As Integer
    
    PosMap As Integer
    PosX As Byte
    PosY As Byte
    PosName As String
    
    Blocked As Byte
    
    Colour As Long

End Type

Public Type tAccount

    SelectedChar As Byte
    Email As String
    DNI As String
    
    Alias As String
    
    Key As String
    Passwd As String
    
    DateBirth As String
    DateRegister As String
    Premium As Byte
    
    CharsAmount As Byte
    Chars(ACCOUNT_MAX_CHARS) As tAccountChar
    LoggedFailed As Byte
    LoggedChar As Boolean
    
    Gld As Long
    Eldhir As Long
    BlockedChars As Byte
    KeyMao As String
    
    ColorNick As Long
    
    RenderHeading As Integer
    SaleSlot As Integer
    
End Type

Public TempAccount          As tAccount

Public Account              As tAccount

Public NullAccount          As tAccount

Public ObjBlacksmith_Amount As Integer

Public ObjBlacksmith()      As tItemsConstruibles

Public ObjBlacksmith_Copy() As tItemsConstruibles

' MAPS
' CONSTANT MAPS
Public Const XMaxMapSize    As Byte = 100 'Map sizes in tiles

Public Const XMinMapSize    As Byte = 1 'Map sizes in tiles

Public Const YMaxMapSize    As Byte = 100 'Map sizes in tiles

Public Const YMinMapSize    As Byte = 1 'Map sizes in tiles

Public Enum eGrhType

    eNone = 0
    eArbol = 1
    
End Enum

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData

    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    GrhType As eGrhType
    
    Speed As Single
    
    src As wGL_Rectangle
    
    Active As Boolean
    MiniMap_color As Long
    Alpha As Boolean

End Type

Public GrhMeditationForm As grh

Public GrhOpciones       As grh

Public Grh_Antorcha      As grh

Public Grh_Criature      As grh

Public Type tMiniMap_Drop

    Name As String
    Amount As Long
    Probability As Byte

End Type

Public Type tMiniMap_Npc

    NpcIndex As Integer
    
    Name As String
    Body As Integer
    Head As Integer
    cant As Integer
    Exp As Long
    Gld As Long
    Eldhir As Long
    
    NroSpells As Byte
    Spells() As String
    
    NroItems As Byte
    NroDrops As Byte
    
    Obj(1 To 30) As tMiniMap_Drop
    Drop(1 To 30) As tMiniMap_Drop
    Hp As Long
    
    MinHit As Integer
    MaxHit As Integer

End Type

Public Type tMinimap

    Loaded As Boolean
    Name As String
    Pk As Boolean
    
    NpcsNum As Byte
    Npcs() As tMiniMap_Npc

    LvlMin As Byte
    LvlMax As Byte
    
    Guild As Byte
    
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    InviSinEfecto As Byte
    CaenItem As Byte
    
    SubMaps As Byte
    Maps() As Integer
    
    ChestLast As Integer
    Chest() As Integer

End Type

Public MapSelected      As Integer

Public MiniMap_Criature As Integer

Public MiniMap_Last     As Integer

Public MiniMap()        As tMinimap

Public Type tInfoObj

    Last As Integer
    LastObj As Integer

End Type

Public InfoObj                     As tInfoObj

Public Const MAX_OBJ               As Integer = 2000

Public Const ESCRITURAS_CLAN       As Integer = 1746

Public Const ACTA_NACIMIENTO       As Integer = 1745

Public Const PENDIENTE_SACRIFICIO  As Integer = 1465

Public Const FRAGMENTO_HIELO  As Integer = 2207

Public Const MANUSCRITO_1  As Integer = 41
Public Const MANUSCRITO_2  As Integer = 1473

Public Const PERLA_FORTUNA_1       As Integer = 1442

Public Const PERLA_FORTUNA_2       As Integer = 1471

Public Const PENDIENT_GROUP        As Integer = 1479

Public Const VaraMataDragonesIndex As Integer = 1037

Public Const EspadaDiablo          As Integer = 1235


Public Const LAUDMAGICO   As Integer = 167
Public Const ANILLOMAGICO As Integer = 168


' Personaje dibujado en el INFOOBJ
Public Type tChar_InfoObj

    Body As Integer
    Head As Integer
    
    Helm As HeadData
    Shield As ShieldAnimData
    Weapon As WeaponAnimData
    
End Type

Public Char_InfoObj As tChar_InfoObj

Public Type tUpgrade

    ObjIndex As Integer
    RequiredCant As Byte
    Required() As Obj
    RequiredPremium As Long ' Fragmentos Premium

End Type

Public Type ObjData_Skills

    Selected As Byte
    Amount As Integer

End Type

' El cofre en si carga la siguiente información.
Public Type tChest

    Probability As Byte
    
    ' Cargamos la lista de drops disponibles. Cada drop a su vez tiene multiples probabilidades (DROP.INI)
    NroDrop As Byte
    Drop() As Integer
    
    RespawnTime As Long
    ClicTime As Long
    ProbClose As Byte
    ProbBreak As Byte

End Type

' Lista de Objetos
Public Type tObjData
    ID As Integer
    
    Name As String
    GrhIndex As Long
    MinDef As Integer
    MaxDef As Integer
    MinHit As Integer
    MaxHit As Integer
    MinDefRM As Integer
    MaxDefRM As Integer
    ObjType As eOBJType
    Anim As Integer
    AnimBajos As Integer
    Proyectil As Integer
    DamageMag As Byte
    ValueDSP As Long
    ValueGLD As Long
    
    NoSeCae As Byte
    Skin As Byte
    GuildLvl As Byte
    
    Points As Long
    TimeWarp As Long
    TimeDuration As Long
    RemoveObj As Byte
    PuedeInsegura As Byte
     
    Tier As Byte
    Color As Long
    LvlMin As Byte
    LvlMax As Byte
    
    Hombre As Byte
    Mujer As Byte
    RazaHumano As Byte
    RazaDrow As Byte
     
    SkillNum As Byte
    Skill() As ObjData_Skills
    SkillsEspecialNum As Byte
    SkillsEspecial() As ObjData_Skills
    
    Upgrade As tUpgrade
    CP_Valid As Boolean
    CP() As Byte
    
    Chest As tChest
    
    VisualSkin As Byte
End Type

Public NumObjDatas As Integer
Public SkinLast As Integer
Public CopyObjs() As tObjData
Public ObjData()   As tObjData

' Lista de Npcs
Public Type tNpcs

    Name As String
    Desc As String
    Body As Integer
    Head As Integer
    
    NpcType As Byte
    Comercia As Byte
    Craft As Byte
    
    PoderAtaque As Integer
    PoderEvasion As Integer
    
    MinHit As Integer
    MaxHit As Integer
    
    Def As Integer
    DefM As Integer
    
    MaxHp As Long
    
    GiveExp As Long
    GiveGld As Long
    
    NroDrops As Byte
    Drop(1 To MAX_INVENTORY_SLOTS) As tDrops
    
    NroItems As Byte
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    
    RespawnTime As Long

End Type

Public NumNpcs   As Integer

Public NpcList() As tNpcs

' Lista de Quests
Public Type tNpc_Quest

    NpcIndex As Integer
    Amount As Currency
    Hp As Long
    AmountFinished As Long
    Color As Long

End Type

Public Type tObj_Quest

    AmountFinished As Long
    ObjIndex As Integer
    Amount As Integer
    Durabilidad As Long
    Color As Long
    View As Boolean

End Type

Public NpcsUser_Quest(1 To MAXUSERQUESTS) As Integer

Public NpcsUser_QuestIndex                As Integer

Public NpcsUser_QuestIndex_Original       As Integer

Public NpcsUser_Selected                  As Byte

Public QuestNum                           As Integer

Public Type tQuest

    QuestEmpezada As Boolean
    
    Name As String
    Desc() As String
    DescFinish As String
    
    Obj As Byte
    Objs() As tObj_Quest
    
    Npc As Byte
    Npcs() As tNpc_Quest
    
    SaleObj As Byte
    SaleObjs() As tObj_Quest
    
    ChestObj As Byte
    ChestObjs() As tObj_Quest
    
    RewardObj As Byte
    RewardObjs() As tObj_Quest
    
    RewardGld As Long
    RewardExp As Long
    
    LastQuest As Byte
    NextQuest As Byte
    
    Remove As Byte
    
    NpcsUser()          As tNpc_Quest
    ObjsUser()          As tObj_Quest
    ObjsSaleUser()          As tObj_Quest
    ObjsChestUser()          As tObj_Quest

End Type

Public NumQuests        As Integer

Public QuestList()      As tQuest

Public QuestNpc()       As Byte

Public QuestLast        As Byte

Public QuestIndex       As Byte

Public QuestObjIndex    As Byte

Public QuestNpcIndex    As Byte

Public NpcName          As String

Public SelectedNpcIndex As Integer ' Criatura seleccionada para ver la info

Public SelectedObjIndex As Integer ' Objecto seleccionado para ver la info

' MAPAAAAAAAAAAAAAAAAAAAAA
Public Enum ePanelMapa

    eDefault = 0 ' Bienvenida al Mapa. Explicación Básica. Datos de Interes.
    eMapInfo = 1 ' Información del Mapa seleccionado.

End Enum
        
Public PanelMapa         As ePanelMapa

Public Const STAT_MAXELV As Byte = 47

Public Type tEstadisticasUsu

    FragsCiu As Long
    FragsCri As Long

    Clase As Byte
    Raza As Byte
    
    Elv As Byte
    Elu As Long
    Exp As Long

    Promedy As Long
    
    Skills(1 To NUMSKILLS) As Byte
    Skills_Valid(1 To NUMSKILLS) As Boolean

End Type

Public Enum eAccount_PanelSelected

    ePrincipal = 0 ' Panel de Conectarse
    ePanelAccount = 1 ' Panel de la Cuenta
    ePanelAccountCharNew = 2 ' Creación de un Nuevo Personaje
    ePanelAccountRecover = 3 ' Recuperación del a Cuenta
    ePanelAccountRegister = 4 ' Registro de una Nueva Cuenta

End Enum

Public Account_PanelSelected As eAccount_PanelSelected

Public Type tGlobalCounters

    StrenghtAndDextery As Integer
    Invisibility As Integer
    Paralized As Integer

End Type

Public GlobalCounters As tGlobalCounters

' Lista Gráfica
Public hlst           As clsGraphicalList

Public RankList       As clsGraphicalList

Public hlstMercader   As clsGraphicalList

Public ScrollArrastrar As Byte

Public LastScroll     As Byte

Public Type tShop

    ID As Integer
    
    Name As String
    Gld As Long
    Dsp As Long
    Desc() As String
    ObjIndex As Integer
    ObjAmount As Integer
    Points As Integer ' Puntos de Torneo que te pide >> A cambio te da Oro/Dsp
    fX As grh

End Type

Public Shop()     As tShop

Public ShopCopy() As tShop

Public ShopLast   As Integer

Public Type ControlInfo_type

    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single

End Type

Public ControlInfos()  As ControlInfo_type

Public Const MAX_TOP   As Byte = 3

Public Const MAX_MONTH As Byte = 12

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

Public Ranking_Month         As Byte

Public Ranking(1 To MAX_TOP) As tRank

' Se podria decir lo nuevo?

Public Type tSkin

    Last As Integer
    ObjIndex() As Integer
        
    ' Actual equipped
    ArmourIndex As Integer
    ShieldIndex As Integer
    WeaponIndex As Integer
    WeaponDagaIndex As Integer
    WeaponArcoIndex As Integer
    HelmIndex As Integer
            
End Type

Public ClientInfo As tUser_General

' BLOQUE PRINCIPAL DE ORGANIZACION NUEVO
Public Type tUser_General
        
    Skin As tSkin                  ' Skins in Objects

End Type

' <<<<<< Targets >>>>>>
Public Enum TargetType

    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
    uArea = 5

End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo

    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3
    uInvocacion = 4

End Enum

Public Type tHechizo

    AutoLanzar As Byte
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    '    Resis As Byte
    
    Tipo As TipoHechizo
    
    TileRange As Byte
    WAV As Integer
    FXgrh As Integer
    Loops As Byte
    
    SubeHP As Byte
    MinHp As Integer
    MaxHp As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    SanacionGlobal As Byte
    SanacionGlobalNpcs As Byte
    
    Warp As Byte
    Invoca As Byte
    NumNpc As Integer
    cant As Integer

    '    Materializa As Byte
    '    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer
    HpRequerido As Integer
    
    AreaX As Byte
    AreaY As Byte
    
    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean

End Type

Public NumeroHechizos As Integer

Public Hechizos()     As tHechizo

Type tLottery
    Name As String
    Desc As String
    
    DateInitial As String ' Inicio del sorteo
    DateFinish As String ' Fecha en la que se realiza el sorteo
    
    PrizeChar As String ' Personaje que va a ser sorteado
    PrizeObj As Integer ' Objeto que va a ser sorteado
    PrizeObjAmount As Integer   ' Cantidad del objeto que recibe
    
    CharLast As Integer
    Chars() As String
    LastSpam As Long                     ' Tiempo que hace que spameo en la consola sobre el sorteo.
End Type

Public Lottery() As tLottery



Public Type Inventario

    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    AuraEqpObjIndex As Integer
    AuraEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MochilaEqpObjIndex As Integer
    MochilaEqpSlot As Byte
    FactionArmourEqpObjIndex As Integer
    FactionArmourEqpSlot As Byte
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    ReliquiaObjIndex As Integer
    ReliquiaSlot As Byte
    PendientePartySlot As Byte
    PendientePartyObjIndex As Integer
    MagicObjIndex As Integer
    MagicSlot As Byte
    
    NroItems As Integer

End Type


' # Carga Bonus1=Tipo|Value|ObjIndex|Amount|DurationUse|DurationDate
Public Type UserBonus
    Tipo As eBonusType
    Value As Long
    Amount As Integer
    
    DurationSeconds As Long         ' Duración en segundos. Esto descuento por uso online del personaje.
    DurationDate As String      ' Duración en fecha (Si termina un dia determinado) (Formato 22-22-2222 18:00:00hs)
    
End Type

Public Type tSkins

    Last As Integer                     ' Numero de Skins que hay en el PJ
    ObjIndex() As Integer           ' Lista de skins aprendidas en el PJ
        
    ArmourIndex As Integer           ' Body asignado actual
    ShieldIndex As Integer             '
    WeaponIndex As Integer
    WeaponDagaIndex As Integer
    WeaponArcoIndex As Integer
    
    HelmIndex As Integer
            
End Type

Public Type tInfoUser
    UserName As String
    Clase As Byte
    Raza As Byte
    Genero As Byte
    Elv As Byte
    Exp As Long
    Elu As Long
    
    Blocked As Byte
    BlockedHasta As Long
    
    Gld As Long
    Dsp As Long
    Points As Long
    Frags As Integer
    FragsCiu As Integer
    FragsCri As Integer
    
    Map As Integer
    X As Byte
    Y As Byte
    
    Hp As Integer
    
    Inventory As Inventario
    Bank As Inventario
    Spells(1 To MAXHECHI) As Integer
    Skills(1 To NUMSKILLS) As Integer
    Skins As tSkins
    
    PenasTime As Integer
    PenasLast As Byte
    Penas() As String
    
    BonusLast As Byte
    BonusUser() As UserBonus
End Type

' # Estadísticas del personaje
Public InfoUser As tInfoUser


' Tipos de Bonus que puedo otorgar
Public Enum eBonusType
    eGLD = 1        ' Altera el % de Oro
    eExp = 2        ' Altera el % de Exp
    ePoints = 3     ' Altera el % de Puntos
    eDSP = 4        ' Altera el % de Dsp
    eObj = 5        ' Determina si da un objeto y no un efecto.
    eVip = 6        ' Agrega tiempo V.I.P
    eMap = 7        ' Agrega acceso a un mapa específico.
End Enum

