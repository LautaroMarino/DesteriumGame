Attribute VB_Name = "Declaraciones"
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

Public FrasesLastMap As Integer
Public FrasesOnFire() As String

Public lastRunTime As String

Public AutoRestart As Boolean

Public Enum e_EstadoMimetismo

    Desactivado = 0
    FormaUsuario = 1
    FormaBichoSinProteccion = 2
    FormaBicho = 3

End Enum

Public Enum e_Facciones

    Armada = 1
    Caos = 2
    Concilio = 3
    Consejo = 4

End Enum

' constantes de sistema de bots inteligente

Public Const BOT_NPCINDEX      As Integer = 175

Public Const BOT_MAX_USER      As Integer = 10       ' Máximo de 10 invocables en caso de que el usuario necesite

Public Const BOT_MAX_SPAWN     As Integer = 500 ' @ A probar performance mama!

Public Const BOT_MAX_INVENTORY As Byte = 20

Public Const BOT_MAX_SPELLS    As Byte = 25

Public Enum eMovementBot

    BOT_MOVEMENT_DEFAULT = 0          ' Aparece y camina de manera random tirando pasos floger
    BOT_GOTOCHAR = 1                         ' Sigue al personaje que lo invocó
    BOT_GOTONPC_RANDOM = 2             ' Sigue una criatura random que se encuentre cerca

End Enum

Public Enum eMovementBotAttack

    BOT_MODE_DEFENSE = 0                   ' El BOT prefiere su defensa y la de sus compañeros
    BOT_MODE_ATTACK = 1                    ' El BOT prefiere el ataque hacia otros
    BOT_MODE_MIXED = 2                       ' El BOT hace un MIXED entre defensa/ataque

End Enum

Public Enum eTypeBar

    eTeleportInvoker = 1
        
End Enum

Public Enum eAnuncios

    A_MISION_COMPLETADA = 1
    
End Enum

Public HelpLast     As Byte

Public HelpLines()  As String

Public EsModoEvento As Byte

Public Enum eHechizosIndex

    eDescarga = 23
    eApocalipsis = 25
    eExplosionAbismal = 62
    eTormenta = 15

End Enum

Public BytesTotal    As Long

Public BytesReceived As Long

Public BytesImage()  As Byte

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public UsersBot                  As Integer

Public HappyHour                 As Boolean

Public PartyTime                 As Boolean

''
' Modulo de declaraciones. Aca hay de todo.
'

Public Event_Time_Auto           As Long

Public Const MAX_KEY             As Integer = 500

Public SecurityKey(0 To MAX_KEY) As Byte

Public MercaderActivate          As Boolean

Public CountDownLimpieza         As Integer

Public GlobalActive              As Boolean

Public CountDown_Time            As Byte

Public CountDown_Map             As Integer

Public ModoNavidad               As Byte

Public MultExp                   As Byte

Public MultGld                   As Byte

Public aClon                     As clsAntiMassClon

Public TrashCollector            As Collection

Public Const MAXSPAWNATTEMPS = 60

Public Const INFINITE_LOOPS       As Integer = -1

''
' The maximum number of private messages.
Public Const MAX_PRIVATE_MESSAGES As Byte = 5

''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL   As Long = &HF82FF

''
' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND          As Byte = 0



    Public Const iFragataFantasmal = 87

    Public Const iFragataReal = 87

    Public Const iFragataCaos = 87

    Public Const iBarca = 84

    Public Const iGalera = 85

    Public Const iGaleon = 86

    Public Const iBarcaCiuda = 84

    Public Const iBarcaReal = 84

    Public Const iBarcaPk = 84

    Public Const iBarcaCaos = 84

    Public Const iGaleraCiuda = 85

    Public Const iGaleraReal = 85

    Public Const iGaleraPk = 85

    Public Const iGaleraCaos = 85

    Public Const iGaleonCiuda = 86

    Public Const iGaleonReal = 86

    Public Const iGaleonPk = 86

    Public Const iGaleonCaos = 86

Public Enum eMenues

    ieComerciante = 1
    ieSacerdote
    ieGobernador
    iemascota
    ieMascotaQuieta
    ieEntrenador
    ieFogata
    ieFogataDescansando
    ieBanquero
    ieEnlistadorFaccion
    ieApostador
    ieYunque
    ieFragua
    ieOtroUser
    ieOtroUserCompartiendoNpc
    ieNpcDomable
    ieLenia
    ieRamas

End Enum

Public Enum iMinerales

    HierroCrudo = 190
    PlataCruda = 191
    OroCrudo = 192
    
    LingoteDeHierro = 193
    LingoteDePlata = 194
    LingoteDeOro = 195

End Enum

Public Enum PlayerType

    User = &H1
    SemiDios = &H2
    LIBRE = &H4
    Dios = &H8
    Admin = &H10
    ChaosCouncil = &H20
    RoyalCouncil = &H40

End Enum

Public Enum ePrivileges

    Admin = 1
    Dios
    Especial
    SemiDios

End Enum

Public Enum eFaccion

    fArmada = 1
    fLegion = 2

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
    cNix = 2
    cBanderbill = 3
    cLindos = 4
    cArghal = 5
    cArkhein = 6
    cNewbie = 7
    cEsperanza = 8
    
    cLastCity

End Enum

Public Enum eRaza

    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano

End Enum

Enum eGenero

    Hombre = 1
    Mujer

End Enum

Public Enum eClanType

    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal

End Enum

Public Const LimiteNewbie As Byte = 12

Public Type tCabecera 'Cabecera de los con

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public MiCabecera                    As tCabecera

Public Const NingunEscudo            As Integer = 2

Public Const NingunCasco             As Integer = 2

Public Const NingunArma              As Integer = 2

Public Const NingunAura              As Integer = 0

Public Const EspadaMataDragonesIndex As Integer = 1037

Public Const PENDIENTE_SACRIFICIO    As Integer = 1465

Public Const ESCRITURAS_CLAN         As Integer = 1746

Public Const ACTA_NACIMIENTO         As Integer = 1745

Public Const VaraMataDragonesIndex   As Integer = 1037

Public Const EspadaDiablo            As Integer = 1235



Public Const LAUDMAGICO   As Integer = 167
Public Const ANILLOMAGICO As Integer = 168

Public Const APOCALIPSIS_SPELL_INDEX As Integer = 25
Public Const DESCARGA_SPELL_INDEX    As Integer = 23

Public Const SLOTS_POR_FILA          As Byte = 5

Public Const PROB_ACUCHILLAR         As Byte = 20

Public Const DAÑO_ACUCHILLAR As Single = 0.2

Public Const MAXMASCOTASENTRENADOR As Byte = 7

'Public Const FXSANGRE = 14

#If Classic = 0 Then

    Public Enum FXIDs

        FXWARP = 42
    
        FXMEDITARCHICO = 4
        FXMEDITARMEDIANO = 5
        FXMEDITARGRANDE = 6
        FXMEDITARXGRANDE = 16
        FXMEDITARXXGRANDE = 29
        FXMEDITARXXXGRANDE = 70
    
        FXSANGRE = 75
        FXSWING = 76
    
        FX_INCINERADO = 72
    
        FX_LEVEL = 96 ' Pasaje de Nivel
        FX_APUÑALADA = 58 ' Efecto al Apuñalar

    End Enum


#Else


    Public Enum FXIDs

        FXWARP = 1
        FXMEDITARCHICO = 4
        FXMEDITARMEDIANO = 5
        FXMEDITARGRANDE = 6
        FXMEDITARXGRANDE = 16
        FXMEDITARXXGRANDE = 16
        FXMEDITARXXXGRANDE = 70
    
        FXSANGRE = 75
        FXSWING = 76
    
        FX_INCINERADO = 72
    
        FX_LEVEL = 96 ' Pasaje de Nivel
        FX_APUÑALADA = 58 ' Efecto al Apuñalar

    End Enum

#End If

Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

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
    zonaOscura = 7
    CASA = 8
    AutoResu = 9
    LavaActiva = 10

End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6

    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3

End Enum

Public Enum eTerrain

    terrain_bosque = 0
    terrain_nieve = 1
    terrain_desierto = 2
    terrain_ciudad = 3
    terrain_campo = 4
    terrain_dungeon = 5

End Enum

Public Enum eRestrict

    restrict_no = 0
    restrict_newbie = 1
    restrict_armada = 2
    restrict_caos = 3
    restrict_faccion = 4

End Enum

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

Public Const MAXUSERHECHIZOS      As Byte = 35

' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 1

Public Const EsfuerzoTalarLeñador As Byte = 1

Public Const EsfuerzoPescarPescador        As Byte = 1

Public Const EsfuerzoPescarGeneral         As Byte = 5

Public Const EsfuerzoExcavarMinero         As Byte = 1

Public Const EsfuerzoExcavarGeneral        As Byte = 5

Public Const FX_TELEPORT_INDEX             As Integer = 1

Public Const PORCENTAJE_MATERIALES_UPGRADE As Single = 0.85

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo

    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6

End Enum

Public Const Guardias                       As Integer = 6

Public Const MAX_ORO_EDIT                   As Long = 50000000

Public Const MAX_VIDA_EDIT                  As Long = 30000

Public Const STANDARD_BOUNTY_HUNTER_MESSAGE As String = "Se te ha otorgado un premio por ayudar al proyecto reportando bugs, el mismo está disponible en tu bóveda."

Public Const TAG_USER_INVISIBLE             As String = "[INVISIBLE]"

Public Const TAG_CONSULT_MODE               As String = "[CONSULTA]"

Public Const TAG_GAME_MASTER                As String = "<Staff>"

Public Const MAXREP                         As Long = 6000000

Public Const MAXORO                         As Long = 2000000000

Public Const MAXEXP                         As Long = 2000083607

Public Const MAXUSERMATADOS                 As Long = 65000

Public Const MAXATRIBUTOS As Byte = 50
Public Const MINATRIBUTOS  As Byte = 6

Public Const LingoteHierro As Integer = 193

Public Const LingotePlata  As Integer = 194

Public Const LingoteOro    As Integer = 195

Public Const Leña As Integer = 196
Public Const LeñaRoble As Integer = 198
Public Const LeñaTejo As Integer = 197

Public Const MAXNPCS  As Integer = 10000

Public Const MAXCHARS As Integer = 10000


Public Const HACHA_LEÑADOR As Integer = 200
Public Const PIQUETE_MINERO As Integer = 199
Public Const CAÑA_PESCA As Integer = 201
Public Const RED_PESCA           As Integer = 399

Public Const DAGA                As Integer = 15

Public Const FOGATA_APAG         As Integer = 136

Public Const FOGATA              As Integer = 63

Public Const ORO_MINA            As Integer = 194

Public Const PLATA_MINA          As Integer = 193

Public Const HIERRO_MINA         As Integer = 192

Public Const MARTILLO_HERRERO    As Integer = 389

Public Const SERRUCHO_CARPINTERO As Integer = 198

Public Const ObjArboles          As Integer = 4
Public Const CAÑA_COFRES As Integer = 1340

Public Enum eNPCType

    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    GuardiasCaos = 8
    ResucitadorNewbie = 9
    Pretoriano = 10
    Gobernador = 11
    Mascota = 12
    Fundition = 13
    eCommerceChar = 14              '  Comerciante que le pertenece a un personaje

End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS         As Byte = 15

''
' Cantidad de skills especiales
Public Const NUMSKILLSESPECIAL As Byte = 9

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS      As Byte = 5

''
' Cantidad de Clases


Public Const NUMCLASES As Byte = 9


' Maximo de Auras sobre el Personaje, ubicadas en lugares estratégicos según el Index.
Public Const MAX_AURAS      As Byte = 5

''
' Cantidad de Razas
Public Const NUMRAZAS       As Byte = 5

''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
' Cantidad de Ciudades
Public Const NUMCIUDADES    As Byte = 8

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

Public Const MAX_QUEST_SIMULTANEO As Integer = 3

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS          As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO             As Integer = 100

Public Const vlASESINO            As Integer = 1000

Public Const vlCAZADOR            As Integer = 5

Public Const vlNoble              As Integer = 5

Public Const vlLadron             As Integer = 25

Public Const vlProleta            As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%





Public Const iORO    As Byte = 12

Public Const iELDHIR As Integer = 943

Public Const Pescado As Byte = 139

#If Classic = 0 Then

    Public Enum PECES_POSIBLES

        PESCADO1 = 326
        PESCADO2 = 327
        PESCADO3 = 328
        PESCADO4 = 329
        PESCADO5 = 330
        PESCADO6 = 331
        PESCADO7 = 332

    End Enum

    Public Const NUM_PECES As Integer = 7

#Else

    Public Enum PECES_POSIBLES

        PESCADO1 = 139
        PESCADO2 = 544
        PESCADO3 = 545
        PESCADO4 = 546
        PESCADO5 = 732

    End Enum

    Public Const NUM_PECES As Integer = 5

#End If

Public ListaPeces(1 To NUM_PECES) As Integer

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
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
    Comerciar = 11
    Domar = 12
    Proyectiles = 13
    Navegacion = 14
    Resistencia = 15

End Enum

Public Enum eMochilas

    Mediana = 1
    Grande = 2

End Enum

Public Const FundirMetal = 88

Public Const TeleportInvoker = 99

Public Enum eAtributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5

End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel

Public Const AdicionalHPCazador  As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef        As Byte = 15

Public Const AumentoStBandido    As Byte = AumentoSTDef + 3

Public Const AumentoSTLadron     As Byte = AumentoSTDef + 3

Public Const AumentoSTMago       As Byte = AumentoSTDef - 1

Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize         As Byte = 100

Public Const XMinMapSize         As Byte = 1

Public Const YMaxMapSize         As Byte = 100

Public Const YMinMapSize         As Byte = 1

'Tamaño del tileset
Public Const TileSizeX           As Byte = 32

Public Const TileSizeY           As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion

Public Const XWindow             As Byte = 17

Public Const YWindow             As Byte = 13

'Sonidos
Public Const SND_SWING           As Byte = 2

Public Const SND_TALAR           As Byte = 13

Public Const SND_PESCAR          As Byte = 14

Public Const SND_MINERO          As Byte = 15

Public Const SND_WARP            As Byte = 3

Public Const SND_PUERTA          As Byte = 5

Public Const SND_NIVEL           As Byte = 6

Public Const SND_USERMUERTE      As Byte = 11

Public Const SND_IMPACTO         As Byte = 10

Public Const SND_IMPACTO2        As Byte = 12

Public Const SND_LEÑADOR As Byte = 13

Public Const SND_FOGATA       As Byte = 14

Public Const SND_AVE          As Byte = 21

Public Const SND_AVE2         As Byte = 22

Public Const SND_AVE3         As Byte = 34

Public Const SND_GRILLO       As Byte = 28

Public Const SND_GRILLO2      As Byte = 29

Public Const SND_SACARARMA    As Byte = 25

Public Const SND_ESCUDO       As Byte = 37

Public Const MARTILLOHERRERO  As Byte = 41

Public Const LABUROCARPINTERO As Byte = 42

Public Const SND_BEBER        As Byte = 46

Public Enum eSound

    sRemove = 16        ' Ruido de remover parálisis
    
    sOlla = 124          ' Sonido de Olla (Pociones) Sonido ambiental mientras se construye
    
    sConstruction = 237 ' Construcción de Crafting criatura
    sConquistCastle = 238 ' Sonido cuando está por conquistar el castillo.
    eDopaPerdida = 239 ' Sonido aleatoreo cuando perdes la dopa
    sVictory = 240 ' Victoria de combate
    sVictory2 = 241 ' Victoria de Combate
    sExplotionAbismal = 242 ' Hechizo Explosion Abismal
    sJuiceFinal = 243 ' Juicio Final. Grito desgarrador
    sVictory3 = 244 ' Victoria de Combate
    sVictory4 = 245 ' Victoria de Combate
    sVictory5 = 246 ' Victoria de Combate
    sIraMelkor = 247 ' Ira de Melkor (Hechizo Gm's)
    sGotmul = 248 ' Grito de Gothul (Dopa Full)
    sSanation = 249 ' Hechizo de Sanacion
    sSanation2 = 250 ' Hechizo de Sanacion n°2
    sSanation3 = 251 ' Hechizo de Sanacion n°3
    sSanation4 = 252 ' Hechizo de Sanacion n°4
    sSanation5 = 253 ' Hechizo de Sanacion n°5
    sSanation6 = 254 ' Hechizo de Sanacion n°6
    sMaterializar = 255 ' Hechizo Materializar
    sAlarmConquist = 256 ' Alarma de Conquista
    sJugosGastricos = 257 ' Hechizo de ataque
    sSpellDefense = 258 ' Hechizo de Defensa (Clerigo)
    sSpellDefense2 = 259 ' Hechizo de Defensa (Clerigo)
    sRafaga = 260 ' Hechizo de Rafaga
    
    sDoubleKill = 261
    sTripleKill = 270
    sUltraKill = 271
    sUnstoppable = 272
    sMonsterKill = 267
    sMegaKill = 267
    sRampage = 269
    sHolyShit = 264
    sPerspal = 262
    sGodlike = 263
    
    sApuñaladaEspalda = 220    ' Ruido de apuñalada por espalda
    sEquipeEspada = 208         ' Sonido al equipar una espada/Daga
    
    sViento = 278   ' Posible viento reto
    sFogata = 411
    sFlechaPega = 412   ' Flecha pegando
    sFlechaFalla = 413    ' Flecha falla
    sIncineracion = 414   ' Efecto de Fuego
    
    sChestBreak = 452       ' El cofre se rompe
    sChestClose = 473       ' El cofre se abre y se vuelve a cerrar
    
    sChestDrop1 = 422
    sChestDrop2 = 447
    sChestDrop3 = 451
    
    sWarp10s = 507
    sWarp20s = 508
    sWarp30s = 509
    sWarp60s = 510

End Enum

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS         As Integer = 10000

''
' Cantidad de Skins por personaje
Public Const MAX_INVENTORY_SKINS        As Byte = 50

''
' Cantidad de "slots" en el inventario con mochila
Public Const MAX_INVENTORY_SLOTS        As Byte = 30

''
' Cantidad de "slots" en el inventario sin mochila
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 30

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO                    As Integer = MAX_INVENTORY_SLOTS + 1

Public Const FLAGELDHIR                 As Integer = FLAGORO + 1

Public Const FLAG_AGUA                  As Byte = &H20

Public Const FLAG_ARBOL                 As Byte = &H40

' CATEGORIAS PRINCIPALES
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
    otItemNpc = 43
    
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

'Texto
Public Const FONTTYPE_TALK            As String = "~255~255~255~0~0"

Public Const FONTTYPE_FIGHT           As String = "~255~0~0~1~0"

Public Const FONTTYPE_WARNING         As String = "~32~51~223~1~1"

Public Const FONTTYPE_INFO            As String = "~65~190~156~0~0"

Public Const FONTTYPE_INFOBOLD        As String = "~65~190~156~1~0"

Public Const FONTTYPE_EJECUCION       As String = "~130~130~130~1~0"

Public Const FONTTYPE_PARTY           As String = "~255~180~255~0~0"

Public Const FONTTYPE_VENENO          As String = "~0~255~0~0~0"

Public Const FONTTYPE_GUILD           As String = "~255~255~255~1~0"

Public Const FONTTYPE_SERVER          As String = "~0~185~0~0~0"

Public Const FONTTYPE_GUILDMSG        As String = "~228~199~27~0~0"

Public Const FONTTYPE_CONSEJO         As String = "~130~130~255~1~0"

Public Const FONTTYPE_CONSEJOCAOS     As String = "~255~60~00~1~0"

Public Const FONTTYPE_CONSEJOVesA     As String = "~0~200~255~1~0"

Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"

Public Const FONTTYPE_CENTINELA       As String = "~0~255~0~1~0"

'Estadisticas
Public Const STAT_MAXELV              As Byte = 47

Public Const STAT_MAXHP               As Integer = 5000

Public Const STAT_MAXSTA              As Integer = 999

Public Const STAT_MAXMAN              As Integer = 9999

Public Const STAT_MAXHIT_UNDER36      As Byte = 99

Public Const STAT_MAXHIT_OVER36       As Integer = 999

Public Const STAT_MAXDEF              As Byte = 99

Public Const ELU_SKILL_INICIAL        As Byte = 250

Public Const EXP_ACIERTO_SKILL        As Byte = 50

Public Const EXP_FALLO_SKILL          As Byte = 50

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************


Public Type tObservacion

    Creador As String
    Fecha As Date
    
    Detalles As String

End Type

Public Type tRecord

    Usuario As String
    Motivo As String
    Creador As String
    Fecha As Date
    
    NumObs As Byte
    Obs() As tObservacion

End Type

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
    loops As Byte
    
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
    
    LvlMin As Byte

End Type

Public Const MAX_PACKET_COUNTERS As Long = 15

Public Enum PacketNames

    CastSpell = 1
    WorkLeftClick
    LeftClick
    UseItem
    UseItemU
    Walk
    sailing
    Talk
    Attack
    Drop
    Work
    EquipItem
    GuildMessage
    QuestionGM
    ChangeHeading

End Enum

Public Type eInfoSkill

    Name As String
    MaxValue As Integer

End Type

Public Type eLevelSkill

    LevelValue As Integer

End Type

Public Type UserOBJ

    ObjIndex As Integer
    Amount As Long
    Equipped As Byte

End Type

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

Public Type Position

    X As Integer
    Y As Integer

End Type

Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

Public Type FXdata

    Nombre As String
    GrhIndex As Long
    Delay As Integer

End Type

Public Type t_Caminata

    offset As Position
    Espera As Long

End Type

'Datos de user o npc
Public Type Char

    charindex As Integer
    Head As Integer
    Body As Integer
    BodyIdle As Integer
    BodyAttack As Integer
    AuraIndex(1 To MAX_AURAS) As Byte
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As eHeading
    speeding As Single

End Type

Public Type tEffects

    Hp As Integer
    Man As Integer
    
    Damage As Integer
    DamageMagic As Integer
    
    DamageNpc As Single
    DamageMagicNpc As Single
    
    AfectaParalisis As Byte
    DevuelveVidaPorc As Byte
    
    ExpNpc As Single

End Type

Public Enum eEffectObj

    e_Gld = 1 ' Da Oro
    e_Exp = 2 ' Da experiencia
    e_Revive = 3 ' Resucita al personaje
    e_NewHead = 4 ' Da una nueva cabeza (V4)
    e_NewHeadClassic = 5 ' Da una cabeza de las clásicas (TDS)
    e_ChangeGenero = 6      ' Cambia el género del personaje

End Enum

Public Type Obj

    ObjIndex As Integer
    Amount As Long

End Type

Public Const FRAGMENTO_PREMIUM As Integer = 1466

Public Type tUpgrade

    ObjIndex As Integer
    RequiredCant As Byte
    Required() As Obj
    RequiredPremium As Long ' Fragmentos Premium

End Type

Public Type tChestObj

    ObjIndex As Integer
    minAmount As Integer
    maxAmount As Integer

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

Public Type ObjData_Skills

    Selected As Byte
    Amount As Integer

End Type

Public Type tColor

    r As Byte
    G As Byte
    B As Byte
    Alpha As Byte

End Type

'Tipos de objetos
Public Type ObjData
    AntiFrio As Byte            ' Sistema de anti Frio en tunicas
    DurationDay As Integer     ' Agrega X cantidad de días
    Dead As Byte           ' Solo se puede usar
    
    
    SkillNum As Byte
    Skill() As ObjData_Skills
    SkillsEspecialNum As Byte
    SkillsEspecial() As ObjData_Skills
    
    MaxFortunas As Byte
    Fortuna() As Obj

    GuildLvl As Byte    ' Requiere un Nivel de Clan
    Skin As Byte
    SizeWidth As Long ' Graphical entity's width
    SizeHeight As Long ' Graphical entity's height
    Navidad As Byte
    DosManos As Byte
    NpcBonusDamage As Integer
    NoShield As Byte
    NoNada As Byte
    NoDrop As Byte
    Plata As Byte
    Bronce As Byte
    Oro As Byte
    Premium As Byte
    
    EdicionLimitada As Byte
    
    MagiaSkill As Byte
    RMSkill As Byte
    ArmaSkill As Byte
    EscudoSkill As Byte
    ArmaduraSkill As Byte
    ArcoSkill As Byte
    DagaSkill As Byte
    QuitaEnergia As Integer
    
    AuraIndex(1 To MAX_AURAS) As Byte
    EffectUser As tEffects
    Name As String 'Nombre del obj
    
    Sound As Integer
    
    Range As Integer
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    Donable As Byte
    
    GrhIndex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    Acuchilla As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHp As Integer ' Minimo puntos de vida
    MaxHp As Integer ' Maximo puntos de vida
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    proyectil As Integer
    Municion As Integer
    Ilimitado As Byte
    
    Crucial As Byte
    Newbie As Integer
    
    MinHitMag As Integer
    MaxHitMag As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    
    Oculto As Byte
    
    LingoteIndex As Integer
    
    MinHit As Integer 'Minimo golpe
    MaxHit As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    RopajeEnano As Integer      'Indice del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    WeaponRazaEnanaAnim As Integer
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    VictimAnim As Integer
    Incineracion As Byte
    
    ValorDefault As Long    ' Precio default en oro
    
    Valor As Long     ' Precio
    ValorEldhir As Long
    Tier As Byte
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    Radio As Integer ' Para teleps: El radio para calcular el random de la pos destino
    
    MochilaType As Byte 'Tipo de Mochila (1 la chica, 2 la grande)
    
    Guante As Byte ' Indica si es un guante o no.
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    Time As Long
    RemoveObj As Byte
    BonusTipe As eEffectObj
    BonusValue As Single
    
    Chest As tChest
    BonoExp As Single
    BonoGld As Single
    BonoEvasion As Single
    BonoRm As Single
    BonoArcos As Single
    BonoArmas As Single
    BonoHechizos As Single
    BonoTime As Integer
    
    GuildExp As Long
    
    TelepMap As Integer
    TelepX As Integer
    TelepY As Integer
    TelepTime As Integer
    RequiredNpc As Integer
    
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    Texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    ProbPesca As Byte
    
    NoSeCae As Integer
    
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
    Points As Long
    Upgrade As tUpgrade
    ArbolItem As Integer
    MenuIndex As Byte

    LvlMin As Byte
    LvlMax As Byte
    
    ' Objetos teleport invoker
    TimeWarp As Long
    TimeDuration As Long
    Position As WorldPos
    TeleportObj As Integer
    PuedeInsegura As Byte
    FX As Integer
    
    VisualSkin As Byte              ' Determina si es cargado por el sistema de skins que visualiza TODOS los items del juego
    Porc As Byte
End Type

'[Pablo ToxicWaste]
Public Type ModClase

    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double

End Type

Public Type ModRaza

    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single

End Type


' BALANCE NUEVA ORGANIZACION

Public Type tBalance_ClassObj
    Obj() As Integer
End Type

Public Type tBalance
    PASSIVE_MAX As Integer
    
    ModClase(1 To NUMCLASES)           As ModClase
    ModRaza(1 To NUMRAZAS)             As ModRaza
    ModVida(1 To NUMCLASES)            As Double
    DistribucionEnteraVida(1 To 5)     As Integer
    DistribucionSemienteraVida(1 To 4) As Integer
    PorcentajeRecuperoMana            As Integer
    
    ListObjs(1 To NUMCLASES) As tBalance_ClassObj
    RazeClass(1 To NUMCLASES) As Byte
    GeneroClass(1 To NUMCLASES) As Byte
    
    Health_Initial(1 To NUMCLASES) As Single
    Health_Level(1 To NUMCLASES) As Single
    
    Mana_Initial(1 To NUMCLASES) As Single
    Mana_Level(1 To NUMCLASES) As Single
    
    Damage_Initial(1 To NUMCLASES) As Single
    Damage_Level(1 To NUMCLASES) As Single
    
    DamageMag_Initial(1 To NUMCLASES) As Single
    DamageMag_Level(1 To NUMCLASES) As Single
    
    Armour_Initial(1 To NUMCLASES) As Single
    Armour_Level(1 To NUMCLASES) As Single
    
    ArmourMag_Initial(1 To NUMCLASES) As Single
    ArmourMag_Level(1 To NUMCLASES) As Single
    
    Attack_Initial(1 To NUMCLASES) As Single
    Attack_Level(1 To NUMCLASES) As Single
    
    RegHP_Initial(1 To NUMCLASES) As Single
    RegHP_Level(1 To NUMCLASES) As Single
    
    RegMANA_Initial(1 To NUMCLASES) As Single
    RegMANA_Level(1 To NUMCLASES) As Single
    
    Movement_Initial(1 To NUMCLASES) As Single
    
    Cooldown_Initial(1 To NUMCLASES) As Single
    Cooldown_Level(1 To NUMCLASES) As Single
End Type

Public Balance As tBalance

''''''''''''''''


'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40

'[/KEVIN]

'[KEVIN]
Public Type BancoInventario

    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer

End Type

'[/KEVIN]

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

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tReputacion 'Fama del usuario

    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    promedio As Long

End Type

Public Type tCharacterStats

    PassiveAccumulated As Long
    
End Type



' Estadisticas de los usuarios
Public Type UserStats

    BonusTipe As eBonusType
    BonusValue As Single
    BonusLast As Integer
    Bonus() As UserBonus
    
    OldHp As Integer
    OldMan As Integer
    Points As Long
    
    'Retos1Jugados As Long
    'Retos1Ganados As Long
    
    'DesafiosJugados As Long
    'DesafiosGanados As Long
    
    'TorneosJugados As Long
    'TorneosGanados As Long
    
    BonosHp As Byte
    Gld As Long 'Dinero
    Eldhir As Long 'Dinero
    

    
    ' MOD TDS
    MinSta As Integer: MaxSta As Integer
    MinHam As Integer: MaxHam As Integer
    MinAGU As Integer: MaxAGU As Integer
    MinHit As Integer: MaxHit As Integer                ' // ADAPTAR por DAMAGE
    '
    
    MinHp As Integer: MaxHp As Integer
    MinMan As Integer: MaxMan As Integer
    
    
    Damage As Single            ' Daño
    DamageMag As Single      ' Daño Mágico
    Armour As Single             ' Armadura
    ArmourMag As Single       ' Armadura Magica
    Attack As Single               ' Velocidad de Ataque
    RegHP As Single               ' Regeneración de Vida
    RegMANA As Single            ' Regeneración de Maná
    Movement As Single          ' Velocidad del Personaje
    Cooldown As Single          ' Cooldown ni me acuerdo que era
    
    Exp As Double
    Elv As Byte
    Elu As Long
    
    UserSkillsEspecial(1 To NUMSKILLSESPECIAL) As Integer
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
    ExpSkills(1 To NUMSKILLS) As Long
    EluSkills(1 To NUMSKILLS) As Long
    
End Type

'Flags
Public Type UserFlags
    Rachas As Integer       ' Rachas que tiene el personaje acumuladas historicamente.
    RachasTemp As Integer   ' Rachas que tiene en el momento para luego chequear por las historicas y reemplazar.
    RedValid As Boolean     ' Está en un evento que pide validación doe rjas.
    RedUsage As Integer     ' Pociones rojas que consumió.
    RedLimit As Integer     ' Limite de RED que tiene impuesto según el evento que está participando.
    BotList As Byte
    TeleportInvoker As Byte
    LastInvoker As Long
    TempAccount As String
    TempPasswd As String
    DeslogeandoCuenta As Boolean
    StreamUrl As String
    ModoStream As Boolean
    Blocked As Byte
    Incinerado As Byte
    ObjIndex As Integer
    ClainObject As Byte
    ToleranceCheat As Byte
    DragBlocked As Boolean
    Transform As Byte
    TransformVIP As Byte
    GmSeguidor As Integer
    MenuCliente As Byte
    LastSlotClient As Byte
    
    DesafiosGanados As Long
    Desafiando As Byte
    SelectedBono As Integer
    
    ' Retos y eventos
    SlotEvent As Byte
    SlotReto As Byte
    SlotRetoUser As Byte
    SlotFast As Byte
    SlotFastUser As Byte
    SlotUserEvent As Byte
    FightTeam As Byte
    SelectedEvent As Byte
    ' /
    
    Plata As Byte
    Bronce As Byte
    Premium As Byte
    Oro As Byte
    Streamer As Byte
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    NoPuedeSerAtacado As Boolean
    AtacablePor As Integer
    ShareNpcWith As Integer
    
    Vuela As Byte
    Navegando As Byte
    Montando As Integer
    Seguro As Boolean
    SeguroResu As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    OwnedNpc As Integer ' Npc que le pertenece (no puede ser atacado)
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    NPCAtacado As Integer
    Ignorado As Boolean
    
    EnConsulta As Boolean
    SendDenounces As Boolean
    
    StatsChanged As Byte
    Privilegios As PlayerType
    PrivEspecial As Boolean
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    
    ChatColor As Long
    
    '[el oso]
    MD5Reportado As String
    '[/el oso]
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    Silenciado As Byte
    
    Mimetizado As Byte
    
    LastMap As Integer
    Traveling As Byte 'Travelin Band ¿?
    
    ParalizedBy As String
    ParalizedByIndex As Integer
    ParalizedByNpcIndex As Integer
    
End Type

Public Type t_ControlHechizos

    HechizosTotales As Long
    HechizosCasteados As Long

End Type

Public Type UserCounters
    CaspeoTime As Long
    BuffoAceleration As Long
    TimerMeditar As Long
    TiempoInicioMeditar As Long
    IntervaloCaminar As Long
    
    TimeGMBOT As Long
    controlHechizos As t_ControlHechizos
    LastSave As Long
    Incinerado As Long
    TimeApparience As Long
    RuidoPocion As Long
    RuidoDopa As Long
    SpamMessage As Long
    MessageSend As Long
    ShieldBlocked As Integer
    Shield As Integer
    ReviveAutomatic As Integer
    FightInvitation As Integer
    FightSend As Integer
    Drawers As Integer
    DrawersCount As Integer
    
    TimeEquipped As Long
    TimeInfoMao As Long
    TimePublicationMao As Integer
    TimerPuedeCastear As Long
    TimerPuedeRecibirAtaqueCriature As Long
    
    TimeDrop As Long
    TimeInfoChar As Long
    TimeMessage As Long
    TimeCommerce As Long
    Packet250 As Long
    Packet500 As Long
    
    TimeInactive As Integer
    TimeCreateChar As Long
    TimeDenounce As Integer
    TimeBonus As Long
    TimeGlobal As Integer
    TimeTransform As Integer
    TimeBono As Integer
    TimeFight As Integer
    TimeTelep As Integer
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    Lava As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    
    Invisibilidad As Integer
    TiempoOculto As Integer
    
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    ' Anticheat
    SpeedHackCounter As Single
    LastStep As Long
    
    TimerShiftear As Long
    'TimerMoverUsuario As Long
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerUsarClick As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long
    TimerEstadoAtacable As Long
    
    Trabajando As Long
    Ocultando As Long
    
    failedUsageAttempts As Long
    failedUsageAttempts_Clic As Long
    failedUsageCastSpell As Long
    
    goHome As Long
    goHomeSec As Long
    AsignedSkills As Byte
    
    TimeLastReset As Long
    PacketCount As Long
    
End Type

Public Type tCrafting

    cantidad As Long
    PorCiclo As Integer

End Type

Public Type tUserMensaje

    Contenido As String
    Nuevo As Boolean

End Type

' Configuración de Personajes de la cuenta
Public Type tCuentaUser

    Name As String
    Ban As Byte
    Clase As eClass
    Raza As eRaza
    Elv As Byte
    
    Body As Integer
    Head As Integer
    Helm As Integer
    Shield As Integer
    Weapon As Integer

End Type

' ///////////////////////////////////////////////////////////////// RETOS
Public Const MAX_RETOS_SIMULTANEOS As Integer = 70

Public Enum eTipoReto

    None = 0
    FightOne = 1
    FightTwo = 2
    FightThree = 3
    FightClan = 4

End Enum

Public Const MAX_RETOS_TERRENOS As Byte = 8

Public Enum eTerreno

    eNormal = 1
    eDungeon = 2
    ePradera = 3
    eNevado = 4
    eMagma = 5
    eMagma1 = 6
    eDesierto = 7
    eNieve = 8

End Enum

Public Const MAX_RETOS_CONFIG As Byte = 6

Public Enum eRetoConfig

    eInmovilizar = 0
    eResucitar = 1
    eEscudos = 2
    eCascos = 3
    eItems = 4
    eFuegoAmigo = 5

End Enum

Private Type tMapEvent

    Run As Boolean
    
    Map As Integer
    X(1) As Byte
    Y(1) As Byte
    
    Zona As eTerreno

End Type

Public Retos(1 To MAX_RETOS_SIMULTANEOS) As tFight

Public Enum eFight

    eNormal = 1
    ePlantes = 2

End Enum

Public Type tFightUser

    Team As Byte
    Name As String
    Accepts As Byte
    UserIndex As Integer
    Rounds As Byte

End Type

Public Type tFight

    Run As Boolean
    
    Tipo As Byte
    Gld As Long
    RoundsLimit As Byte
    Time As Long
    Terreno As eTerreno
    User() As tFightUser
    config(1 To MAX_RETOS_CONFIG) As Byte
    RoundsGanados As Byte
    
    TimeSound As Long
    Arena As Integer
    
    ' Datos temporales para dejar tiempo de descansos
    TimeDescanso As Long
    TeamDescanso As Byte
    ChangeRound As Boolean

End Type

'############################################### FIN RETOS
                            
' Quests / Misiones
Public Type tQuestNpc

    NpcIndex As Integer
    Amount As Long
    Hp As Long

End Type
 
Public Type tUserQuest
    
    ObjsPick() As Long
    ObjsSale() As Long
    NPCsKilled() As Long
    QuestIndex As Integer

End Type

Public DailyLast    As Byte

Public QuestDaily() As Byte

Public Type tQuest

    Nombre As String
    Desc As String
    DescFinish As String
    RequiredLevel As Byte
    RequiredBronce As Byte
    RequiredPlata As Byte
    RequiredOro As Byte
    RequiredPremium As Byte
    
    RequiredOBJs As Byte
    RequiredObj() As Obj
    
    RequiredSaleOBJs As Byte
    RequiredSaleObj() As Obj
    
    RequiredChestOBJs As Byte
    RequiredChestObj() As Obj
    
    RequiredNPCs As Byte
    RequiredNpc() As tQuestNpc
   
    RewardGLD As Long
    RewardEldhir As Long
    RewardEXP As Long
   
    RewardOBJs As Byte
    RewardObj() As Obj
    
    DoneQuest As Byte
    DoneQuestMessage As String
    
    LastQuest As Byte
    NextQuest As Byte
    
    RewardDaily As Byte '   La Quest es Diaria

    Remove As Byte

End Type

' End Quest / Misiones

Public Type tMascotasUser

    Name As String
    Elv As Byte
    Elu As Long
    Exp As Long
    MinHp As Integer
    MaxHp As Integer
    MinMan As Integer
    MaxMan As Integer
    MinHit As Integer
    MaxHit As Integer
    MinHitMag As Integer
    MaxHitMag As Integer
    ClassValid(1 To NUMCLASES) As Byte
    RazeValid(1 To NUMRAZAS) As Byte
    Spells(1 To 35) As Integer

End Type

' Control de Anti Frags
Public Const MAX_CONTROL_FRAGS As Byte = 15

Public Type tAntiFrags

    UserName As String
    Account As String
    IP As String
    
    Time As Long
    cant As Byte

End Type

' Almacenamos la información del personaje
Public Type tOldInfoUser

    MaxHp As Integer
    MaxMan As Integer
    Clase As eClass
    Raza As eRaza
    UserSpell(1 To MAXUSERHECHIZOS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserSkillsEspecial(1 To NUMSKILLS) As Integer
    UserSkills(1 To NUMSKILLS) As Byte
    Items(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    GldRed As Long
    GldBlue As Long
    MinHit As Integer
    MaxHit As Integer
    MaxSta As Integer
    
    Elv As Byte
    Exp As Long
    
    Head As Integer

End Type

Public Type tUserSecurity

    IUse As Long
    IClick As Long
    IDrop As Long
    ISpell As Long
    IAttack As Long
    ILeftClick As Long

End Type

' CUENTAS

' Subasta de Objetos
Public Type tAccountMercader_Obj

    ObjData As Obj
    Gld As Long
    Eldhir As Long
    Time As Long

End Type

' Subasta de Personajes
Public Type tAccountMercader_Char

    Char As String
    Gld As Long
    Eldhir As Long
    Time As Long
    
End Type

Public Type tAccountChar

    Slot As Byte
    Name As String
    Blocked As Byte
    Guild As String
    
    Map As Integer
    
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
    posX As Integer
    posY As Integer

End Type

Public Type tAccountSecurity

    IP_Server As String
    IP_Public As String
    IP_Local As String
    IP_Address As String
    
    SERIAL_MAC As String
    SERIAL_DISK As String
    SERIAL_MOTHERBOARD As String
    SERIAL_BIOS As String
    SERIAL_PROCESSOR As String
     
    SYSTEM_DATA As String

End Type

Public NullAccount As tAccount

Public Type tAccount

    SlotLogged As Byte ' Pj LOGEADO
    MercaderSlot As Integer
    Sec As tAccountSecurity
    Alias As String
    Email As String
    
    FirstName As String
    LastName As String
    
    Key As String
    KeyMao As String
    Passwd As String
    
    DateBirth As String
    DateRegister As String
    DatePremium As String
    Premium As Byte
    
    CharsAmount As Integer
    Chars(ACCOUNT_MAX_CHARS) As tAccountChar
    LoggedFailed As Byte
    
    'AuctionObj(1 To ACCOUNT_MAX_AUCTION_OBJ) As tAccountMercader_Obj
    'AuctionChar As tAccountMercader_Char
    
    BancoInvent As BancoInventario
    
    Gld As Long
    Eldhir As Long

End Type

Public Enum eLastPotion

    eNullPotion = 0
    eRed = 1
    eBlue = 2

End Enum

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

' SISTEMA DE BOTS INTELIGENTES
Public Type tBotIntelligence_Config
        
    InventoryInitial() As Integer
    SpellsInitial() As Integer

End Type

Public Type tBotIntelligence_Stats

    Elv As Byte
    Exp As Long
    Elu As Long
        
    MinHp As Long
    MaxHp As Long
        
    MinMan As Long
    MaxMan As Long
        
    MinHit As Long
    MaxaHit As Long

End Type

Public Type tBotIntelligence

    Active As Boolean
    
    Class As eClass
    Raze As eRaza
    Name As String
    Head As Integer
    
    Inventory() As Obj
    
    Spells() As Integer
    Skins As tSkins
    Movement As eMovementBot
    MovementAttack As eMovementBotAttack
    
    Stats As tBotIntelligence_Stats
    
    WeaponIndex As Integer
    ArmourIndex As Integer
    HelmIndex As Integer
    ShieldIndex As Integer

End Type

Public BotIntelligence_Config(1 To NUMCLASES) As tBotIntelligence_Config

Public BotIntelligence()                      As tBotIntelligence

'Tipo de los Usuarios
Public Type User

    ' NUEVA ORGANIZACIÓN
    CharacterStats As tCharacterStats
    
    ' FIN NUEVA ORG
    RankMonth As Byte
    LastPotion As eLastPotion
    PotionBlue_Clic As Byte
    PotionBlue_Clic_Interval As Byte
    
    PotionBlue_U As Byte
    PotionBlue_U_Interval As Byte
    
    PotionRed_Clic As Byte
    PotionRed_Clic_Interval As Byte
    
    PotionRed_U As Byte
    PotionRed_U_Interval As Byte
    
    UseObj_Clic As Byte
    UseObj_Init_Clic As Long
    UseObj_U As Byte
    UseObj_Init_U As Long
    Next_UseItem As Boolean
    
    InfoMap As Boolean
    AccountLogged As Boolean
    
    Skins As tSkins
    LastClick As WorldPos
    LastClickCant As Long
    ObjectClaim(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    
    UserLastClick As Long
    UserLastClick_Tolerance As Byte
    UserKey As Long
    DañoApu As Long
    OldInfo As tOldInfoUser
    MeditationSelected As Byte
    MeditationUser(0 To MAX_MEDITATION) As Integer
        
    AntiFrags(MAX_CONTROL_FRAGS) As tAntiFrags
    KeyPackets(MAX_KEY_PACKETS) As tPackets
    Pointers(1 To MAX_POINTERS) As tPoint
    interval(1) As tUserSecurity
    
    InfoMascotas As tMascotasUser
    RetoTemp As tFight
    PosAnt As WorldPos
    IsPremium As Boolean
    Name As String
    
    secName As String
    ID As Long
    
    ShowName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    Char As Char 'Define la apariencia
    CharMimetizado As Char
    OrigChar As Char
    
    Desc As String ' Descripcion
    DescRM As String
    
    Blocked As Byte
    BlockedHasta As Long
    
    Clase As eClass
    Raza As eRaza
    Genero As eGenero
    Hogar As eCiudad
        
    Invent As Inventario
    
    Pos As WorldPos
    Power As Boolean
    
    ConnIDValida As Boolean
    LastPacket As Boolean
    LastPacketError As Long
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    
    PenasLast As Byte
    Penas() As String
    Construir As tCrafting
    
    MascotaIndex As Integer
    
    ServerSelected As Byte
    Stats As UserStats
    flags As UserFlags
    LastHeading As Byte
    
    Reputacion As tReputacion
    
    Faction As tFaction

    #If ConUpTime Then
        LogOnTime As Date
        UpTime As Long
    #End If

    IpAddress As String
    
    ComUsu As tCOmercioUsuario
    
    GuildIndex As Integer   'puntero al array global de guilds
    GuildRange As eGuildRange
    GuildSlot As Byte
    EscucheClan As Integer
    
    GroupIndex As Integer   'index a la party q es miembro
    GroupRequest As String ' Personaje que le ofreció pertenecer a la Party
    GroupRequestTime As Long ' Dilay para evitar flodeos
    GroupSlotUser As Byte ' Slot de miembro
    
    KeyCrypt As Integer
    
    CurrentInventorySlots As Byte
    
    QuestLast As Integer
    QuestStats() As tUserQuest
    
    Mensajes(1 To MAX_PRIVATE_MESSAGES) As tUserMensaje
    UltimoMensaje As Byte

    Account As tAccount
    
    LastRequestLogin As Long
    
    PosOculto As WorldPos
    
    MacroIterations(1 To MAX_PACKET_COUNTERS) As Long
    PacketTimers(1 To MAX_PACKET_COUNTERS) As Long
    PacketCounters(1 To MAX_PACKET_COUNTERS) As Long
    
    BotIntelligence(1 To BOT_MAX_USER) As tBotIntelligence
    
    TimeUseClicInitial As Long
    TimeUseClic As Long
End Type

Public MacroIterations(1 To MAX_PACKET_COUNTERS)      As Long

Public PacketTimerThreshold(1 To MAX_PACKET_COUNTERS) As Long

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats

    MaxHp As Long
    MinHp As Long
    MaxMan As Long
    MinMan As Long
    MaxHit As Integer
    MinHit As Integer
    def As Integer
    defM As Integer
    Elv As Byte

End Type

Public Type NpcCounters
    
    Incinerado As Long
    Paralisis As Integer
    TiempoExistencia As Long
    Velocity As Long
    UseItem As Long
    Attack As Long
    Descanso As Long
    MovimientoConstante As Long
    RuidoPocion As Long

End Type

Public Enum e_Alineacion

    ninguna = 0
    Real = 1
    Caos = 2

End Enum

Public Type NPCFlags

    NpcIdle As Boolean
    Invasion As Byte
    RespawnTime As Long
    Invocation As Byte
    TeamEvent As Byte
    InscribedPrevio As Byte
    SlotEvent As Byte
    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    
    ExpGuildCount As Long
    ExpCount As Long
    ResourceCount As Long
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    KeepHeading As Byte
    Sound As Integer
    AttackedBy As String
    AttackedByInteger As Integer
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    RespawnOrigPosRandom As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Incinerado As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    AtacaUsuarios As Boolean ' Si el NPC puede atacar usuarios
    AtacaNPCs As Boolean     ' Si el NPC puede atacar otros NPC
    AIAlineacion As e_Alineacion

End Type

Public Type tCriaturasEntrenador

    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer

End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo

    PathLength As Integer   ' Number of steps *
    Path() As tVertice      ' This array holds the path
    Destination As Position ' The location where the NPC has to go
    RangoVision As Single
    Inteligencia As Integer
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type

' New type for holding the pathfinding info

Public Type tDrops

    ObjIndex As Integer
    Amount As Long
    Probability As Byte
    ProbNum As Byte

End Type

Public Const MAX_NPC_DROPS As Byte = 5

' # INTELIGENCIA PROPIA. CARACTERISTICAS BY ARGENTUM GAME
Public Enum eBot_Action

    ACTION_AFK = 0 ' La criatura está quieta. Esperando una orden
    ACTION_MOVEMENT = 1 'La criatura está yendo del Punto A => Punto B
    ACTION_RETURN = 2   'La Criatura está yendo del Punto B => Punto A
    ACTION_ATTACK = 3  'La Criatura comienza a atacar

End Enum

Public Enum eResourceType

    eMineral = 1
    eLeña = 2
    ePeces = 3
        
End Enum

Public Enum ePretorianAI

    King = 1        ' Rey Supremo
    SpellCaster = 2     ' Lanzador de Hechizos
    SwordMaster = 3  ' Lanzador de Golpes
    MixCaster = 4       ' Lanzador de Golpes y Hechizos
    ArrowCaster = 5   ' Lanzador de Flechas
    
    Last

End Enum

Public Type Npc
    CastleIndex As Integer
    CommerceChar As String          ' El comerciante pertenece a algun personaje
    CommerceIndex As Integer        ' Indice para identificar cual mercader es
    
    BotIndex As Long                    ' Index de BOT
    PretorianAI As ePretorianAI     ' Establece la inteligencia del equipo BOSS
    Quest As Byte
    Quests() As Byte
    Action As eBot_Action
    
    PosA As WorldPos
    PosB As WorldPos
    PosC As WorldPos

    Level As Byte
    ShowName As Byte
    MonturaIndex As Integer
    
    Name As String
    TempDrops As String
    Char As Char 'Define como se vera
    Desc As String

    NPCtype As eNPCType
    numero As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte
    
    LastHeading As eHeading
    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer
    
    Velocity As Integer
    IntervalAttack As Integer
    
    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Owner As Integer
    QuestNumber As Byte
    
    Distancia As Long
    
    GiveEXPGuild As Long
    GiveEXP As Long
    GiveGLD As Long
    RequiredWeapon As Integer
    AntiMagia As Byte
    
    GiveResource As Obj
    
    NroDrops As Byte
    Drop(1 To MAX_INVENTORY_SLOTS) As tDrops
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    Entrenable As Boolean
    MaestroNpc As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    pathFindingInfo As NpcPathFindingInfo
    
    'Hogar
    Ciudad As Byte
    
    MenuIndex As Byte
    
    'Para diferenciar entre clanes
    ClanIndex As Integer
    
    Caminata() As t_Caminata
    CaminataActual As Byte
    
    SizeWidth As Long ' Graphical entity's width
    SizeHeight As Long ' Graphical entity's height
    
    EventIndex As Byte

End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock

    Blocked As Byte
    Graphic(1 To 4) As Long
    UserIndex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
    trigger As eTrigger
    
    ' Para eventos
    ObjEvent As Byte
    
    Protect As Long
    
    TimeClic As Long ' Tiempo que necesito para poder volver a realizar intento de apertura
    TeleportIndex As Long

End Type

'Spawn Pos
Public Type tNpcSpawnPos

    Pos() As Position

End Type

'Info del mapa
Type MapInfo
    OnFire As Byte
    FreeAttack As Boolean
    CanTravel As Boolean
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    OnDeathGoTo As WorldPos
    OnLoginGoTo As WorldPos
    GoToOns As WorldPos
    
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    Exp As Single
    
    ' Anti Magias/Habilidades
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    MimetismoSinEfecto As Byte
    
    RoboNpcsPermitido As Byte
    
    LvlMin As Byte
    LvlMax As Byte
    Premium As Byte
    
    Limpieza As Byte
    CaenItems As Byte
    
    Bronce As Byte
    Plata As Byte
    Guild As Byte
    
    Terreno As String
    Zona As String
    Restringir As Byte
    BackUp As Byte
    NpcSpawnPos(0 To 1) As tNpcSpawnPos
    
    NoMana As Byte
    Players As Network.Group
    
    SubMaps As Byte
    Maps() As Integer
    
    Pesca As Byte
    PescaItem() As Integer
    
    Faction As Byte
    
    
    AccessDays() As Byte ' 1=Lunes, 2=Martes, ..., 7=Domingo
    AccessTimeStarts() As Integer ' Representa HHMM, por ejemplo, 1700 para las 17:00
    accessTimeEnds() As Integer ' Representa HHMM, por ejemplo, 2300 para las 23:00
    
    MinOns As Integer
    
    UsersDead As Integer
    DeadTime As Long
    
    
    Poder As Byte               ' # Mapas en los que funciona el poder
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE                      As Boolean

Public ULTIMAVERSION                     As String

Public BackUp                            As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS)         As String

Public ListaClases(1 To NUMCLASES)       As String

Public ListaAtributos(1 To NUMATRIBUTOS) As String

Public RECORDusuarios                    As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath                           As String

''
'Ruta base para guardar los chars
Public CharPath                          As String

''
'Ruta base para guardar las account
Public AccountPath                       As String

''
'Ruta base para guardar los logs
Public LogPath                           As String

''
'Ruta base para los archivos de mapas
Public MapPath                           As String

''
'Ruta base para los DATs
Public DatPath                           As String

''
'Bordes del mapa
Public MinXBorder                        As Byte

Public MaxXBorder                        As Byte

Public MinYBorder                        As Byte

Public MaxYBorder                        As Byte

''
'Numero de usuarios actual
Public NumUsers                          As Integer

Public LastUser                          As Integer

Public LastChar                          As Integer

Public NumChars                          As Integer

Public LastNPC                           As Integer

Public NumNpcs                           As Integer

Public NumFX                             As Integer

Public NumMaps                           As Integer

Public NumObjDatas                       As Integer

Public NumeroHechizos                    As Integer

Public AllowMultiLogins                  As Byte

Public IdleLimit                         As Integer

Public MaxUsers                          As Integer

Public HideMe                            As Byte

Public LastBackup                        As String

Public Minutos                           As String

Public haciendoBK                        As Boolean

Public PuedeCrearPersonajes              As Integer

Public PuedeConectarPersonajes           As Integer

Public ServerSoloGMs                     As Integer

Public ValidacionDePjs                   As Integer

Public Type tConfigServer

    ModoRetos As Byte
    ModoRetosFast As Byte
    ModoInvocaciones As Byte
    ModoCastillo As Byte
    ModoCrafting As Byte
    ModoSubastas As Byte
    ModoSkins As Byte
End Type

Public ConfigServer                              As tConfigServer

Public NumRecords                                As Integer

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado                       As Byte

Public EnPausa                                   As Boolean

Public EnTesting                                 As Boolean

'*****************ARRAYS PUBLICOS*************************
Public UserList()                                As User 'USUARIOS

Public Npclist(1 To MAXNPCS)                     As Npc 'NPCS

Public MapData()                                 As MapBlock

Public MapInfo()                                 As MapInfo

Public Hechizos()                                As tHechizo

Public CharList(1 To MAXCHARS)                   As Integer

Public ObjData()                                 As ObjData

Public FX()                                      As FXdata

Public SpawnList()                               As tCriaturasEntrenador

Public LevelSkill(1 To 50)                       As eLevelSkill

Public InfoSkill(1 To NUMSKILLS)                 As eInfoSkill

Public InfoSkillEspecial(1 To NUMSKILLSESPECIAL) As eInfoSkill

Public ForbidenNames()                           As String

Public ForbidenText()                            As String

Public BanIps                                    As Collection

Public Ciudades(1 To NUMCIUDADES)                As WorldPos

Public distanceToCities()                        As HomeDistance

Public Records()                                 As tRecord

Public QuestList()                               As tQuest
'*********************************************************

Type HomeDistance

    distanceToCity(1 To NUMCIUDADES) As Integer

End Type

Public Nix          As WorldPos

Public Ullathorpe   As WorldPos

Public Banderbill   As WorldPos

Public Lindos       As WorldPos

Public Arghal       As WorldPos

Public Esperanza    As WorldPos

Public Newbie       As WorldPos

Public CiudadFlotante       As WorldPos

Public Arkhein      As WorldPos

Public Nemahuak     As WorldPos

Public Prision      As WorldPos

Public Libertad     As WorldPos

Public Denuncias    As cCola

Public SonidosMapas As SoundMapInfo

Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpFileName As String) As Long

Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (ByRef Destination As Any, _
                                      ByVal Length As Long)

Public Enum e_ObjetosCriticos

    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467

End Enum

Public Enum eMessages

    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldother
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
    Home
    CancelHome
    FinishHome
    DragSafeOn
    DragSafeOff
    ModoStreamOn
    ModoStreamOff

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
    ChangeMapExp              '/MODMAPINFO EXP 1.1
    ChangeMapInfoAttack    '/MODMAPINFO ATAQUE 1
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

Public Const MATRIX_INITIAL_MAP As Integer = 1

Public Const GOHOME_PENALTY     As Integer = 5

Public Const GM_MAP             As Integer = 49

#If Classic = 0 Then

    Public Const TELEP_OBJ_INDEX As Integer = 143

#Else

    Public Const TELEP_OBJ_INDEX As Integer = 378

#End If
#If Classic = 1 Then

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

#Else

    ' Cabezas y Cuerpos versión TDS
    Public Const HUMANO_H_PRIMER_CABEZA  As Integer = 1

    Public Const HUMANO_H_ULTIMA_CABEZA  As Integer = 25

    Public Const HUMANO_H_CUERPO_DESNUDO As Integer = 21

    Public Const ELFO_H_PRIMER_CABEZA    As Integer = 102

    Public Const ELFO_H_ULTIMA_CABEZA    As Integer = 111

    Public Const ELFO_H_CUERPO_DESNUDO   As Integer = 21

    Public Const DROW_H_PRIMER_CABEZA    As Integer = 201

    Public Const DROW_H_ULTIMA_CABEZA    As Integer = 205

    Public Const DROW_H_CUERPO_DESNUDO   As Integer = 32

    Public Const ENANO_H_PRIMER_CABEZA   As Integer = 301

    Public Const ENANO_H_ULTIMA_CABEZA   As Integer = 305

    Public Const ENANO_H_CUERPO_DESNUDO  As Integer = 53

    Public Const GNOMO_H_PRIMER_CABEZA   As Integer = 401

    Public Const GNOMO_H_ULTIMA_CABEZA   As Integer = 405

    Public Const GNOMO_H_CUERPO_DESNUDO  As Integer = 53

    '**************************************************
    Public Const HUMANO_M_PRIMER_CABEZA  As Integer = 71

    Public Const HUMANO_M_ULTIMA_CABEZA  As Integer = 75

    Public Const HUMANO_M_CUERPO_DESNUDO As Integer = 39

    Public Const ELFO_M_PRIMER_CABEZA    As Integer = 170

    Public Const ELFO_M_ULTIMA_CABEZA    As Integer = 176

    Public Const ELFO_M_CUERPO_DESNUDO   As Integer = 39

    Public Const DROW_M_PRIMER_CABEZA    As Integer = 270

    Public Const DROW_M_ULTIMA_CABEZA    As Integer = 276

    Public Const DROW_M_CUERPO_DESNUDO   As Integer = 40

    Public Const ENANO_M_PRIMER_CABEZA   As Integer = 370

    Public Const ENANO_M_ULTIMA_CABEZA   As Integer = 371

    Public Const ENANO_M_CUERPO_DESNUDO  As Integer = 60

    Public Const GNOMO_M_PRIMER_CABEZA   As Integer = 471

    Public Const GNOMO_M_ULTIMA_CABEZA   As Integer = 475

    Public Const GNOMO_M_CUERPO_DESNUDO  As Integer = 60

#End If

' Por ahora la dejo constante.. SI se quisiera extender la propiedad de paralziar, se podria hacer
' una nueva variable en el dat.
Public Const GUANTE_HURTO                           As Integer = 1767

Public Const ESPADA_VIKINGA                         As Integer = 123

'''''''
'' Pretorianos
'''''''
Public ClanPretoriano()                             As clsClanPretoriano

Public Const MAX_DENOUNCES                          As Integer = 20

'Mensajes de los NPCs enlistadores (Nobles):
Public Const MENSAJE_REY_CAOS                       As String = "¿Esperabas pasar desapercibido, intruso? Los servidores del Demonio no son bienvenidos, ¡Guardias, a él!"

Public Const MENSAJE_REY_CRIMINAL_NOENLISTABLE      As String = "Tus pecados son grandes, pero aún así puedes redimirte. El pasado deja huellas, pero aún puedes limpiar tu alma."

Public Const MENSAJE_REY_CRIMINAL_ENLISTABLE        As String = "Limpia tu reputación y paga por los delitos cometidos. Un miembro de la Armada Real debe tener un comportamiento ejemplar."

Public Const MENSAJE_DEMONIO_REAL                   As String = "Lacayo de Tancredo, ve y dile a tu gente que nadie pisará estas tierras si no se arrodilla ante mi."

Public Const MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE As String = "Tu indecisión te ha condenado a una vida sin sentido, aún tienes elección... Pero ten mucho cuidado, mis hordas nunca descansan."

Public Const MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE   As String = "Siento el miedo por tus venas. Deja de ser escoria y únete a mis filas, sabrás que es el mejor camino."

Public Administradores                              As clsIniManager

Public Type tRangeGM

    Name As String
    Tag As String

End Type

Public RangeGm()                   As tRangeGM

'Modificador de defensa para armaduras de segunda jerarquía.
Public Const MOD_DEF_SEG_JERARQUIA As Single = 1.25

Public AnimHogar(1 To 4)           As Integer

Public AnimHogarNavegando(1 To 4)  As Integer

'Caracteres
Public Const car_Especiales = "áàäâÁÀÄÂéèëêÉÈËÊíìïîÍÌÏÎóòöôÓÒÖÔúùüûÚÙÜÛñÑ.,;!¿?()-_"

'ELU CARGADO DESDE .INI
Public EluUser(1 To STAT_MAXELV) As Long

Public Enum eDamageType

    d_DañoUser = 1
    d_DañoUserSpell = 2
    d_DañoNpc = 3
    d_DañoNpcSpell = 4
    d_CurarSpell = 5
    d_Apuñalar = 6
    d_DiamontRed = 7
    d_DiamontBlue = 8
    d_AddMan = 9
    d_Exp = 10
    d_AddExpBonus = 11
    d_AddGld = 12
    d_AddGldBonus = 13
    d_Aniquilado = 14
    d_AniquiladoPor = 15
    d_AddMagicWord = 16
    d_DañoNpc_Critical = 17
    d_Fallas = 18

End Enum

Public Const POCION_ROJA As Byte = 38

Public Enum eMoveType

    Inventory = 1
    Bank

End Enum

' Criaturas respawn
Public Const RESPAWN_MAX As Byte = 50

Private Type tRespawnnpc

    Time As Long
    NpcIndex As Integer
    OrigPos As WorldPos
    Map As Integer
    CastleIndex As Integer

End Type

Public Respawn_Npc(1 To RESPAWN_MAX) As tRespawnnpc

Public Type tInvasionNpc

    ID As Integer
    cant As Byte
    Map As Integer

End Type

Public Type tInvasion

    Run As Boolean
    Name As String
    Desc As String
    Duration As Long
    
    Maps() As Integer
    
    Npcs As Integer
    Npc() As tInvasionNpc

    InitialMap As Integer
    InitialX As Integer
    InitialY As Integer
    
    Time As Long

End Type

Public Invations()    As tInvasion

Public Invations_Last As Integer

Public Type tNpc

    NpcIndex As Integer
    Name As String
    
    Body As Integer
    Head As Integer
    cant As Integer
    Exp As Long
    Gld As Long
    Eldhir As Long
    
    NroSpells As Byte
    Spells() As Integer
    
    NroItems As Byte
    Invent As Inventario
    
    NroDrops As Byte
    Drop(1 To MAX_INVENTORY_SLOTS) As tDrops
    Hp As Long
    
    MinHit As Integer
    MaxHit As Integer

End Type

Public Type tMinimap

    Name As String
    Pk As Byte

    Gld As Long
    Eldhir As Long
    
    NpcsNum As Byte
    Npcs(50) As tNpc

    LvlMin As Byte
    LvlMax As Byte
    
    Guild As Byte
    
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    InviSinEfecto As Byte
    CaenItem As Byte
    
    Sub_Maps As Byte
    Maps() As Integer
    
    ChestLast As Integer
    Chest() As Integer

End Type

Public MiniMap()             As tMinimap

Public Const VelocidadNormal As Single = 1

Public Const VelocidadMuerto As Single = 1.4



Public Function iCuerpoMuerto(ByVal Criminal As Boolean)
    
    If Not Criminal Then
        iCuerpoMuerto = 8
    Else
        iCuerpoMuerto = 145
    End If

End Function
Public Function iCabezaMuerto(ByVal Criminal As Boolean)
    
    If Not Criminal Then
        iCabezaMuerto = 500
    Else
        iCabezaMuerto = 501
    End If

End Function
