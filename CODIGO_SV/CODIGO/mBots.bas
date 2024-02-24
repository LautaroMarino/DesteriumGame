Attribute VB_Name = "mBots"
Option Explicit

' 26/02/2023
' Sistema de BOTS: Criaturas inteligentes con posibilidad de adquirir diferentes habilidades y características unicas que lo destaquen frente a las criaturas de otros personajes.

' #USO PRINCIPAL
' Las criaturas son destinadas a la lucha CvC, entendiendo así como C= Criatura, quedando Criatura vs Criatura.
' En un tablero estilo Ajedrez, pero que en vez de tener las piezas al principio, podremos optar por moverlas según nuestra conveniencia, mejorando así la probabilidad de ganar la partida y llevarse el punto.

' #USO NRO°2
' Las criaturas podrán tener una inteligencia controlada por el personaje que desee traerla al mundo. Este caso sería utilizado para luchar contra BOTS INTELIGENTES, en retos.

' #USO NRO°3
' Opción de auto completar con bots inteligentes. Permite así completar grandes eventos, o bien cuando un usuario deslogea en evento, un BOT lo sustituye.

Public Const SPELL_PARALISIS As Byte = 9
Public Const SPELL_REMOVERPARALISIS As Byte = 10
Public Const SPELL_INMOVILIZAR As Byte = 24
Public Const SPELL_DESCARGAELECTRICA As Byte = 23
Public Const SPELL_APOCALIPSIS As Byte = 25

' @ Estos valores podrían ser alterados por NIVEL/CONFIGURACION de las criaturas
Public Const BOT_POCION_ROJA_REGENERATE As Byte = 30
Public Const BOT_POCION_AZUL_REGENERATE As Long = 30                ' * lvl

' @ Lo normal es poner 100. Cada vez que se pone 100+ es un rango con 0.00001
Public Const BOT_MANCO_USEITEM As Long = 200
Public Const BOT_MANCO_ATTACK As Long = 200
Public Const BOT_MANCO_DEFENSE As Long = 200
Public Const BOT_MANCO_CAMINAR As Long = 100

' Cada mejora de npc son 10%+
' Total de Niveles del BOT=10
'TOTAL: 100 DE PROB

Public Function BotIntelligence_Balance_Prob(ByVal Elv As Byte) As Boolean
    
    Dim Temp As Long
    Const PORC_ADD As Byte = 10
    
    Temp = (Elv * (PORC_ADD))

    BotIntelligence_Balance_Prob = (RandomNumber(1, 100) <= Temp)
End Function

' Partiendo del Intervalo de 1.000MS, va descontando según el NIVEL del USUARIO.
Public Function BotIntelligence_Balance_UseItem(ByVal Elv As Byte) As Boolean
    
    Dim Temp As Long
    Const PORC_ADD As Integer = 20
    Const ONE_SECOND As Integer = 1000
    
    BotIntelligence_Balance_UseItem = ONE_SECOND - (Elv * (PORC_ADD + 10))
    
End Function

' Posibilidad de manquear según el NIVEL
Public Function BotIntelligence_Balance_Prob_Manco(ByVal Elv As Byte, ByVal max As Long) As Boolean
    Dim Temp As Long
    
    Const PORC_ADD As Byte = 10
    
    Temp = (Elv * (PORC_ADD))

    BotIntelligence_Balance_Prob_Manco = (RandomNumber(1, max) <= Temp)
End Function
' Cada Nivel parte de una base multiplicado por el nivel mismo.
Private Function BotIntelligence_ELU(ByVal Elv As Byte) As Long

    Const BASE_ADD As Long = 1000000
    BotIntelligence_ELU = Elv * BASE_ADD
End Function

' @ Cargamos la configuración inicial del BOT
Public Sub BotIntelligence_Load()

    Dim A         As Long, B As Long
    
    Dim Text      As String

    Dim TextVal() As String

    Dim FilePath  As String
    
    FilePath = DatPath & "bots.ini"

    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    Manager.Initialize FilePath
    
    ' @ Inventarios Iniciales
    For A = 1 To NUMCLASES

        With BotIntelligence_Config(A)
        
            ' Inventory
            Text = Manager.GetValue("INVENTORY", ListaClases(A))
            TextVal = Split(Text, "-")
            
            If UBound(TextVal) <> -1 Then
                ReDim .InventoryInitial(LBound(TextVal) To UBound(TextVal))
            
                For B = LBound(TextVal) To UBound(TextVal)
                    .InventoryInitial(B) = val(TextVal(B))
                Next B
            
            End If
            
            ' Spells
            Text = Manager.GetValue("SPELLS", ListaClases(A))
            TextVal = Split(Text, "-")
            
            If UBound(TextVal) <> -1 Then
                ReDim .SpellsInitial(LBound(TextVal) To UBound(TextVal))
            
                For B = LBound(TextVal) To UBound(TextVal)
                    .SpellsInitial(B) = val(TextVal(B))
                Next B

            End If

        End With
    
    Next A

    ReDim BotIntelligence(1 To BOT_MAX_SPAWN) As tBotIntelligence
    Set Manager = Nothing

End Sub

' @ Activamos un bot al personaje con opción de Spawn
Public Function BotIntelligence_Add(ByVal UserIndex As Integer, ByVal Name As String, ByVal Class As eClass, ByVal Raze As eRaza, ByVal Head As Integer, ByVal Spawn As Boolean) As Boolean
        
    Dim A        As Long

    Dim NpcIndex As Integer

    Dim Pos      As WorldPos

    Dim cChar    As Char
    
    Dim Slot     As Integer
    
    Slot = BotIntelligence_SlotUser(UserIndex)
    
    If Slot = 0 Then
        Call WriteConsoleMsg(UserIndex, "¡No tienes mas lugar para agregar mascotas!", FontTypeNames.FONTTYPE_INFORED)
        Exit Function

    End If
    
    With UserList(UserIndex).BotIntelligence(Slot)
        .Name = Name
        .Class = Class
        .Raze = Raze
        .Head = Head
        
        ' Comienza con atributos INICIALES.
        .Stats.Elv = val(frmMain.txtMascota.Text)
        .Stats.Exp = 0
        .Stats.Elu = BotIntelligence_ELU(.Stats.Elv)
        .Stats.MaxHp = 340                  ' BALANCE DE CLASES-RAZA NORMAL
        .Stats.MaxMan = 2420             ' BALANCE DE CLASES-RAZA NORMAL
        
        ' @ La clase elegida tiene inventario por defecto.
        If UBound(BotIntelligence_Config(.Class).InventoryInitial) <> -1 Then
                
            ReDim .Inventory(1 To BOT_MAX_INVENTORY) As Obj
                
            For A = 1 To UBound(BotIntelligence_Config(.Class).InventoryInitial) + 1
                .Inventory(A).ObjIndex = BotIntelligence_Config(.Class).InventoryInitial(A - 1)
                .Inventory(A).Amount = 1
            Next A

        End If
            
        ' @ La clase elegida tiene Spells por defecto.
        If UBound(BotIntelligence_Config(.Class).SpellsInitial) <> -1 Then
                
            ReDim .Spells(1 To BOT_MAX_SPELLS) As Integer
                
            For A = 1 To UBound(BotIntelligence_Config(.Class).SpellsInitial) + 1
                .Spells(A) = BotIntelligence_Config(.Class).SpellsInitial(A - 1)
            Next A

        End If

    End With
    
    If Spawn Then
        Call BotIntelligence_Spawn(UserList(UserIndex).BotIntelligence(Slot), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, BOT_MOVEMENT_DEFAULT, BOT_MODE_MIXED)

    End If

    BotIntelligence_Add = True

End Function

' @ Spawn de un BOT en una posición elegida.
Public Function BotIntelligence_Spawn(ByRef BotCopy As tBotIntelligence, _
                                      ByVal Map As Integer, _
                                      ByVal X As Byte, _
                                      ByVal Y As Byte, _
                                      ByRef Movement As eMovementBot, _
                                      ByRef MovementAttack As eMovementBotAttack) As Boolean

    Dim Slot     As Long

    Dim NpcIndex As Integer

    Dim Pos      As WorldPos

    Dim cChar    As Char

    Dim A        As Long
    
    Slot = BotIntelligence_FreeSlot(): If Slot = 0 Then Exit Function
    Pos.Map = Map: Pos.X = X: Pos.Y = Y
    NpcIndex = SpawnNpc(BOT_NPCINDEX, Pos, False, False): If NpcIndex = 0 Then Exit Function

    BotIntelligence(Slot) = BotCopy
    
    With BotIntelligence(Slot)
        .Active = True
        .Movement = Movement
        .MovementAttack = MovementAttack
    
        cChar.Head = .Head
        cChar.WeaponAnim = NingunArma
        cChar.CascoAnim = NingunCasco
        cChar.ShieldAnim = NingunEscudo
        cChar.charindex = Npclist(NpcIndex).Char.charindex
        
        'ReDim .Inventory(LBound(BotCopy.Inventory) To UBound(BotCopy.Inventory)) As Obj
        'ReDim .Spells(LBound(BotCopy.Spells) To UBound(BotCopy.Spells)) As Integer
        
        If LBound(.Inventory) Then

            For A = LBound(BotCopy.Inventory) To UBound(BotCopy.Inventory)
                '.Inventory(A) = BotCopy.Inventory(A)
                
                If .Inventory(A).ObjIndex > 0 Then

                    Select Case ObjData(.Inventory(A).ObjIndex).OBJType
                    
                        Case eOBJType.otarmadura
                            cChar.Body = GetArmourAnim_Bot(Slot, .Inventory(A).ObjIndex)
                            .ArmourIndex = .Inventory(A).ObjIndex
                            
                        Case eOBJType.otWeapon
                            cChar.WeaponAnim = GetWeaponAnimBot(.Raze, .Inventory(A).ObjIndex)
                            .WeaponIndex = .Inventory(A).ObjIndex
                            
                        Case eOBJType.otcasco
                            cChar.CascoAnim = ObjData(.Inventory(A).ObjIndex).CascoAnim
                            .HelmIndex = .Inventory(A).ObjIndex
                            
                        Case eOBJType.otescudo
                            cChar.ShieldAnim = ObjData(.Inventory(A).ObjIndex).ShieldAnim
                            .ShieldIndex = .Inventory(A).ObjIndex
                    End Select
                
                End If
                
            Next A

        End If
            
        If UBound(BotCopy.Spells) Then
                
            For A = LBound(BotCopy.Spells) To UBound(BotCopy.Spells)
                .Spells(A) = BotCopy.Spells(A)
            Next A

        End If

    End With
        
    With Npclist(NpcIndex)
        .BotIndex = Slot
        .Stats.Elv = BotCopy.Stats.Elv
        
        .Name = BotCopy.Name
        .Stats.MaxHp = BotCopy.Stats.MaxHp
        .Stats.MaxMan = BotCopy.Stats.MaxMan
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMan = .Stats.MaxMan
        
        .Char = cChar ' @ Toma la apariencia final de la criatura
        
        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, eHeading.SOUTH)

    End With
        
    BotIntelligence_Spawn = True

End Function

' @ Inteligencia de las Criaturas: CAMINATA, ATAQUES
Public Sub BotIntelligence_AI(ByVal NpcIndex As Integer)
    
    Dim CanAttack  As Boolean

    Dim RandomMove As Byte
    
    
    With BotIntelligence(Npclist(NpcIndex).BotIndex)
        
        ' @@ El chamigo ya empezo manqueando, se cae a pedazos.
        'If Not BotIntelligence_Balance_Prob_Manco(.Stats.Elv, BOT_MANCO_CAMINAR) Then Exit Sub
        
        ' @@ Ataques del BOT
        If Intervalo_CriatureAttack(NpcIndex, False) Then
            ' @@ El chamigo manquea el ataque.
            If BotIntelligence_Balance_Prob_Manco(.Stats.Elv, BOT_MANCO_ATTACK) Then
        
            Select Case .MovementAttack
            
                Case eMovementBotAttack.BOT_MODE_ATTACK
                
                Case eMovementBotAttack.BOT_MODE_DEFENSE
                    Call BotIntelligence_CheckEffects(NpcIndex)                 ' Comprueba los Efectos de él (Parálisis,).
                    
                Case eMovementBotAttack.BOT_MODE_MIXED
                    
                    ' 50% de prob de que priorice defensa y luego intente atacar en caso de no haber defendido.
                    If RandomNumber(1, 100) <= 25 Then
                        If Not BotIntelligence_CheckEffects(NpcIndex) Then
                            Call BotIntelligence_AttackAI(NpcIndex)
                        Else
                            Call BotIntelligence_AttackAI(NpcIndex)
                        End If
    
                    Else
                        Call BotIntelligence_AttackAI(NpcIndex)
                    End If

            End Select
            
            End If
        End If
        
        
        ' @ Movimientos del BOT
        If Npclist(NpcIndex).flags.Paralizado + Npclist(NpcIndex).flags.Inmovilizado = 0 Then
            If Intervalo_CriatureVelocity(NpcIndex) Then
                ' @@ El chamigo manquea las teclas a lo loco
             '   If Not BotIntelligence_Balance_Prob_Manco(.Stats.Elv, 50) Then Exit Sub
                
                Select Case .Movement
        
                        ' @ Movimientos aleatoreos por todo el mapa
                    Case eMovementBot.BOT_MOVEMENT_DEFAULT
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    
                        ' @ Movimientos aleatoreos siguiendo a un personaje
                    Case eMovementBot.BOT_GOTOCHAR
                    
                        ' @ Movimientos aleatoreos siguiendo a una criatura
                    Case eMovementBot.BOT_GOTONPC_RANDOM
                        
                End Select
            End If
        End If
        
        
        ' @ Uso de Objetos (HP,MAN)
        If BotIntelligence_Balance_Prob_Manco(.Stats.Elv, 1000) Then
        Call BotIntelligence_RegenerateStats(NpcIndex)
        End If
    End With
  
End Sub

' @ Chequea estados del BOT (Parálisis, invisibilidad, estupidez, incinerado, envenenado, congelado, frizeado)
Public Function BotIntelligence_CheckEffects(ByVal NpcIndex As Integer) As Boolean
    
    Dim CanRemoveParalisis As Boolean

    With Npclist(NpcIndex)
        ' @@ El chamigo manquea el ataque.
        'If Not BotIntelligence_Balance_Prob_Manco(.Stats.Elv, BOT_MANCO_DEFENSE) Then Exit Function
            
        ' ESTADO: Parálisis/Inmovilidad
        If .flags.Paralizado + .flags.Inmovilizado > 0 Then
            If BotIntelligence_Balance_Prob(.Stats.Elv) Then
                CanRemoveParalisis = True
            End If
        End If

        If CanRemoveParalisis Then
            
            ' @ Tiene la MANA suficiente para removerse la parálisis.
            If .Stats.MinMan >= Hechizos(SPELL_REMOVERPARALISIS).ManaRequerido Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SPELL_REMOVERPARALISIS, True)
                .Stats.MinMan = .Stats.MinMan - Hechizos(SPELL_REMOVERPARALISIS).ManaRequerido

            End If
            
            BotIntelligence_CheckEffects = True
            Exit Function
        End If
        
        
    End With

End Function

' @ Comprueba la VIDA/MANA y utiliza Pociones en caso de ser necesario
Private Sub BotIntelligence_RegenerateStats(ByVal NpcIndex As Integer)
    
    Dim RegenerateHP  As Long

    Dim RegenerateMAN As Long
    
    With Npclist(NpcIndex)
        
        ' @@ El chamigo manquea el ataque.
       ' If Not BotIntelligence_Balance_Prob_Manco(.Stats.Elv, BOT_MANCO_USEITEM) Then Exit Sub
            
        ' @ No tienes nada que hacer aquí hombre
        If .Stats.MinHp = .Stats.MaxHp And .Stats.MinMan = .Stats.MaxMan Then Exit Sub
        
        ' @ Intervalo de regeneración según el NIVEL del BOT
        If Not Intervalo_BotUseItem(NpcIndex) Then Exit Sub
        
        ' 7% de prob de no hacer nada por manco
        If RandomNumber(1, 100) <= 7 Then Exit Sub
        
        ' 14% de prob de ir a rojas porque si
        If RandomNumber(1, 100) <= 14 Then
            Call BotIntelligence_EffectPocion(NpcIndex, True)
            Exit Sub
        End If
    
        ' ¿Tiene más del 70% de la vida? ¡PROB DE PASAR A AZULES!
        If .Stats.MinHp >= .Stats.MaxHp * 0.7 Then
            Call BotIntelligence_EffectPocion(NpcIndex, False)
            Exit Sub

        End If
            
        If .Stats.MinHp <> .Stats.MaxHp Then
            Call BotIntelligence_EffectPocion(NpcIndex, True)
            Exit Sub
        End If

    End With

End Sub

' @ Realiza el efecto de la POCIÓN ROJA y POCIÓN AZUL
Private Sub BotIntelligence_EffectPocion(ByVal NpcIndex As Integer, _
                                         ByVal PocionRoja As Boolean)
    
    Dim TempTick As Double

    TempTick = GetTime
    
    With Npclist(NpcIndex)
        
        If PocionRoja Then
            .Stats.MinHp = .Stats.MaxHp + BOT_POCION_ROJA_REGENERATE
            
            If .Stats.MinHp >= .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        Else
            .Stats.MinMan = .Stats.MinMan + CLng(BOT_POCION_AZUL_REGENERATE * .Stats.Elv)

            If .Stats.MinMan >= .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan

        End If
        
        If TempTick - .Contadores.RuidoPocion > 1000 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
            .Contadores.RuidoPocion = TempTick

        End If

    End With
        
End Sub

' @ Busca en el rango de visión criaturas/usuarios posibles de atacar
Private Sub BotIntelligence_AttackAI(ByVal NpcIndex As Integer)
    
    Dim X        As Long, Y       As Long

    Dim NI       As Integer, UI      As Integer

    Dim bNoEsta  As Boolean
    
    Dim SignoNS  As Integer, SignoEO As Integer

    Dim tHeading As Byte
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_x Step IIf(SignoEO = 0, 1, SignoEO)

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        If RandomNumber(1, 100) <= 30 Then
                            NI = MapData(.Pos.Map, X, Y).NpcIndex
                        Else
                            UI = MapData(.Pos.Map, X, Y).UserIndex

                        End If
                        
                        If NI > 0 And NI <> NpcIndex Then                              ' SEARCHED NPCINDEX
                            .TargetNPC = NI
                            
                            '  25% de prob de no querer atacar al NPC de esa fila (porque está paralizado, tampoco la pavada)
                            If RandomNumber(1, 100) <= 25 Then
                                Exit For
                            Else
                                Call BotIntelligence_AttackNPC(NpcIndex, NI)
                                Exit Sub
                            End If

                            Exit Sub
                        
                        ElseIf UI > 0 Then                          ' SEARCHED USERINDEX

                        End If

                    End If

                Next X
            Next Y

        Else

            For Y = .Pos.Y - RANGO_VISION_y To .Pos.Y + RANGO_VISION_y
                For X = .Pos.X - RANGO_VISION_y To .Pos.X + RANGO_VISION_y

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NpcIndex

                        If NI > 0 And NI <> NpcIndex Then
                            
                            ' ¿ Ya tiene TARGET? 70% PROB DE DARLE A ESE CHANGO !
                            If .TargetNPC > 0 Then
                                 If RandomNumber(1, 100) <= 90 Then
                                    
                                    If Distance(Npclist(.TargetNPC).Pos.X, Npclist(.TargetNPC).Pos.Y, .Pos.X, .Pos.Y) <= RANGO_VISION_x Then
                                        Call BotIntelligence_AttackNPC(NpcIndex, .TargetNPC)
                                        Exit Sub
                                    End If
                                 Else
                                    Call BotIntelligence_AttackNPC(NpcIndex, NI)
                                    Exit Sub
                                End If
                            
                            
                            
                            End If
                            
                             .TargetNPC = NI
                        End If
                        
                    End If

                Next X
            Next Y

        End If

    End With

End Sub

' @ Las criaturas van a realizar ataques respecto al 'ARMA' que tengan equipada (Daga,Espada,Vara,Arco o Arrojadizos)
Private Sub BotIntelligence_AttackNPC(ByVal NpcIndex As Integer, ByVal tNpcIndex As Integer)
    
    Dim WeaponIndex As Integer
    
    WeaponIndex = BotIntelligence(Npclist(NpcIndex).BotIndex).WeaponIndex
    
    If WeaponIndex > 0 Then
        If ObjData(WeaponIndex).proyectil = 1 Then
             Call BotIntelligence_ProyectilNPC(NpcIndex, tNpcIndex)
             Exit Sub
             
        ElseIf ObjData(WeaponIndex).StaffDamageBonus <> 0 Then
            Call BotIntelligence_SpellNPC(NpcIndex, tNpcIndex)
            
            Exit Sub
        
        Else
            Call BotIntelligence_CuerpoNPC(NpcIndex, tNpcIndex)
        End If
    
    End If
    
    With Npclist(tNpcIndex)
        If .Stats.MinHp <= 0 Then
            Call MuereNpc(tNpcIndex, 0)
        End If
        
    End With
End Sub

Private Function BotIntelligence_ProyectilNPC(ByVal NpcIndex As Integer, tNpcIndex As Integer) As Boolean
    
End Function

Public Sub BotIntelligence_SpellNPC(ByVal NpcIndex As Integer, ByVal tNpcIndex As Integer)
    
    Dim MageParalisis As Byte
    
    With Npclist(NpcIndex)
        
        ' 55% de que los MAGOS no presten atención a PARALIZAR a las VICTIMAS. (Por anti petes)
        If BotIntelligence(.BotIndex).Class = eClass.Mage Then MageParalisis = 55

        ' 18% de prob DE NO FIJARSE. BONUS MAGO = 55%
        If Not RandomNumber(1, 100) <= (18 + MageParalisis) Then
            If Npclist(tNpcIndex).flags.Paralizado + Npclist(tNpcIndex).flags.Inmovilizado = 0 Then
                If RandomNumber(1, 100) <= 10 Then
                    If BotIntelligence_Balance_Prob(.Stats.Elv) Then
                        If .Stats.MinMan >= Hechizos(SPELL_PARALISIS).ManaRequerido Then
                            Call NpcLanzaSpellSobreNpc(NpcIndex, tNpcIndex, SPELL_PARALISIS, True)
                            .Stats.MinMan = .Stats.MinMan - Hechizos(SPELL_PARALISIS).ManaRequerido
                            Exit Sub

                        End If

                    End If

                    Exit Sub

                End If

            End If

        End If
        
        ' 14% de no tirar ningún hechizo.
        If RandomNumber(1, 100) <= 14 Then Exit Sub
         
        ' Se fija si puede tirar APOCALIPSIS
        If BotIntelligence_Balance_Prob(.Stats.Elv) Then
            If .Stats.MinMan >= Hechizos(SPELL_APOCALIPSIS).ManaRequerido Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, tNpcIndex, SPELL_APOCALIPSIS, True)
                .Stats.MinMan = .Stats.MinMan - Hechizos(SPELL_APOCALIPSIS).ManaRequerido
                Exit Sub
    
            End If
        End If
        
        
        ' 40% de probabilidad de pasar a azules para tirar apocasosssssssssssssssss
        If RandomNumber(1, 100) <= 40 Then
            Call BotIntelligence_EffectPocion(NpcIndex, False)
            Exit Sub
        End If
        
        ' Se fija si puede tirar DESCARGA ELECTRICA
        If BotIntelligence_Balance_Prob(.Stats.Elv) Then
            If .Stats.MinMan >= Hechizos(SPELL_DESCARGAELECTRICA).ManaRequerido Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, tNpcIndex, SPELL_DESCARGAELECTRICA, True)
                .Stats.MinMan = .Stats.MinMan - Hechizos(SPELL_DESCARGAELECTRICA).ManaRequerido
                Exit Sub
    
            End If
        End If
    End With
    
End Sub

Private Sub BotIntelligence_CuerpoNPC(ByVal NpcIndex As Integer, _
                                      ByVal tNpcIndex As Integer)

    With Npclist(NpcIndex)

        If Distancia(.Pos, Npclist(tNpcIndex).Pos) <= 1 Then
            Call SistemaCombate.NpcAtacaNpc(NpcIndex, tNpcIndex)
        End If
    
    End With

End Sub














' # No me interesa verlo





' @ Slot LIBRE para invocar una nueva criatura BOT
Private Function BotIntelligence_FreeSlot() As Long
    Dim A As Long
    
    For A = 1 To BOT_MAX_SPAWN
        With BotIntelligence(A)
            If .Active = False Then
                BotIntelligence_FreeSlot = A
                Exit Function
            End If
        End With
    
    Next A
End Function

' @ Reset información del BOT
Public Sub BotIntelligence_Reset(ByVal Slot As Long)
    Dim NullBot As tBotIntelligence
    
    BotIntelligence(Slot) = NullBot
End Sub

' Buscamos un Slot libre para que le usuario ponga una nueva mascota
Public Function BotIntelligence_SlotUser(ByVal UserIndex As Integer) As Integer
    
    Dim A As Long
    
    With UserList(UserIndex)
    
        For A = 1 To BOT_MAX_USER

            If .BotIntelligence(A).Class = 0 Then
                BotIntelligence_SlotUser = A
                Exit Function

            End If

        Next A
    
    End With

End Function
