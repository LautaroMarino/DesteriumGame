Attribute VB_Name = "EventosDS"
Option Explicit

Public TEMP_SLOTEVENT As Byte ' Info del NPC SELECCIONADO
Public Const MAX_REWARD_OBJ As Byte = 10

Public Enum eModalityEvent

    CastleMode = 1
    DagaRusa = 2
    DeathMatch = 3
    Enfrentamientos = 4
    Teleports = 5
    GranBestia = 6
    Busqueda = 7
    Unstoppable = 8
    JuegosDelHambre = 9
    
    Manual = 10

    
End Enum

Public Enum eConfigEvent
    eBronce = 0
    ePlata = 1
    eOro = 2
    ePremium = 3
    eDañoZona = 4
    eAutoCupos = 5
    eInvFree = 6
    eParty = 7
    eGuild = 8
    eResu = 9
    eOcultar = 10
    eInvisibilidad = 11
    eInvocar = 12
    eMezclarApariencias = 13
    eDagaMaster = 14
    eSpellCuration = 15
    eUsePotion = 16
    eUseParalizar = 17
    eUseApocalipsis = 18
    eUseDescarga = 19
    eUseTormenta = 20
    eTeletransportacion = 21
    eCascoEscudo = 22
    eFuegoAmigo = 23
End Enum

Public Const MAX_EVENTS_CONFIG As Byte = 24

Public Type tEvents
    Predeterminado As Boolean ' Determina si es un evento cargado desde un .ini, de manera tal que se tenga que reingresar al sistema y hacerse de forma ilimitada.
    TimeInit_Default As Long
    TimeCancel_Default As Long

    LastReward As Byte ' Ultimo objeto cargado
    RewardObj(1 To MAX_REWARD_OBJ) As Obj ' Lista de Premios Donados
    Name As String
    Config(0 To MAX_EVENTS_CONFIG - 1) As Byte
    Enabled As Boolean
    Run As Boolean
    Modality As eModalityEvent
    IsPlante As Byte
    TeamCant As Byte

    ArenasOcupadas As Byte
    ArenasLimit As Byte
    ArenasMin As Byte
    ArenasMax As Byte
    
    GanaSigue As Byte   ' Determina a cuantas victorias puede retirarse
    Quotas As Byte
    QuotasMin As Byte
    QuotasMax As Byte
    Inscribed As Byte
    Rounds As Byte
    
    LvlMax As Byte
    LvlMin As Byte
    
    InscriptionGld As Long
    InscriptionGldPremium As Long
    
    AllowedClasses() As Byte
    
    InscriptionAcumulated As Byte
    PrizePoints As Integer
    PrizeExp As Long
    PrizeGldPremium As Integer
    AllowedFaction() As Byte
    PrizeGld As Long
    PrizeObj As Obj
     LimitRed As Long
    TempAdd As String
    TempDate As String
    TempFormat As String
    
    LimitRound As Byte
    LimitRoundFinal As Byte
    
    TimeInscription As Long
    TimeCancel As Long
    TimeCount As Long
    TimeFinish As Long
    TimeInit As Long
    
    ' Por si alguno es con NPC
    NpcIndex As Integer
    
    ' Por si cambia el body del personaje y saca todo lo otro.
    CharHp As Integer
    
    npcUserIndex As Integer
    
    Prob As Byte            ' Si el evento o criatura del evento tiene una PROB se usa esto.
    
    ' Forma la visual del personaje
    iBody As Integer
    iHead As Integer
    iHelm As HeadData
    iShield As ShieldAnimData
    iWeapon As WeaponAnimData
    
End Type

' # Configuración General de los Eventos
Public Const MAX_EVENT_SIMULTANEO As Byte = 20
Public Events(1 To MAX_EVENT_SIMULTANEO) As tEvents
Public Events_TimeInit As Long
Public LastEvent As Byte

Public Function strModality(ByVal Modality As eModalityEvent) As String

10  Select Case Modality

        Case eModalityEvent.CastleMode
20          strModality = "REYVSREY"
                  
30      Case eModalityEvent.DagaRusa
40          strModality = "DAGARUSA"
                  
50      Case eModalityEvent.DeathMatch
60          strModality = "DEATHMATCH"
                  
        Case eModalityEvent.Enfrentamientos
            strModality = "DUELOS"
        
        Case eModalityEvent.Enfrentamientos
            strModality = "MANUAL"
        
        Case eModalityEvent.Teleports
            strModality = "TELEPORTS"
        
        Case eModalityEvent.GranBestia
            strModality = "GRANBESTIA"
        
        Case eModalityEvent.JuegosDelHambre
            strModality = "JUEGOSDELHAMBRE"
            
70          ' Case eModalityEvent.Aracnus
80          'strModality = "Aracnus"
                  
90          ' Case eModalityEvent.HombreLobo
100             'strModality = "HombreLobo"
                  
110             'Case eModalityEvent.Minotauro
120             'strModality = "Minotauro"
                  
130             Case eModalityEvent.Busqueda
140             strModality = "BUSQUEDA"
                  
150            Case eModalityEvent.Unstoppable
160             strModality = "IMPARABLE"
              
170             'Case eModalityEvent.Invasion
180             'strModality = "Invasion"
              
210     End Select

End Function

Public Function CheckAllowedClasses(ByRef AllowedClasses() As Byte) As String

    Dim LoopC As Integer

    Dim Valid As Boolean: Valid = True
    
    For LoopC = 1 To NUMCLASES

        If AllowedClasses(LoopC) = 1 Then
            If CheckAllowedClasses = vbNullString Then
                CheckAllowedClasses = ListaClases(LoopC)
            Else
                CheckAllowedClasses = CheckAllowedClasses & ", " & ListaClases(LoopC)
            End If

        Else
            Valid = False
        End If

    Next LoopC

    If Valid Then
        CheckAllowedClasses = "TODAS"
    End If

End Function

Public Function Events_GetDescription(ByRef Modality As eModalityEvent) As String
    Select Case Modality
    
        Case eModalityEvent.Enfrentamientos
            Events_GetDescription = "Los usuarios van pasando por arenas de combate hasta llegar a la final. Los combates pueden ser 1vs1,2vs2,3vs3,4vs4 y 5vs5."
        
        Case eModalityEvent.CastleMode
            Events_GetDescription = "Los usuarios defienden a dos Reyes en un combate cuerpo a cuerpo. Deberás hacer una estrategía con tu equipo, donde una parte tendrá que defender a vuestro rey y la otra id a atacar al otro. ¡Mucha Suerte! Aaa y casi me olvidaba, al final deberás traicionar a tus compañeros dandole el último golpe al Rey, así te llevarás su Corona y Vestimenta ¡HA HA HA!"
        
        Case eModalityEvent.DagaRusa
            Events_GetDescription = "Un asesino acabará con todos los participantes apuñalandolos por donde mas duela! Aquel suertudo será ganador y el asesino volverá al infierno."
        
        Case eModalityEvent.Teleports
            Events_GetDescription = "Un Teleport te llevará al siguiente nivel. Si la suerte te acompaña, iluminarás el camino de otros y serás o no el ganador final. Recuerda los teleports anteriores para ser el más rapido."
        
        Case eModalityEvent.Unstoppable
            Events_GetDescription = ""
            
        Case Else
            Events_GetDescription = "Descripción no disponible..."
    
    End Select

End Function

Public Sub Events_GenerateSpam(ByVal Slot As Byte, ByRef Console As RichTextBox)

    Dim strTemp  As String
    Dim txtRojas As String
    
    ' Define RGB colors for different types of text
    Dim TitleColor() As Variant: TitleColor = Array(50, 205, 50)    ' Green
    Dim DescColor() As Variant: DescColor = Array(70, 130, 180) ' Steel Blue
    Dim ValueColor() As Variant: ValueColor = Array(255, 165, 0) ' Orange

    With Events(Slot)
        ' Modality-Rounds
        If (.Modality = Enfrentamientos) Then
            strTemp = " | Rounds: " & .LimitRound & IIf(.LimitRound > 1, "s", vbNullString) & IIf(.LimitRoundFinal <> .LimitRound, ". (Final a " & .LimitRoundFinal & ")", vbNullString)
        End If

        ' Title
        AddtoRichTextBox Console, "'" & UCase$(.Name) & "'" & IIf(.Config(eConfigEvent.eFuegoAmigo) = 1, " (Fuego Amigo)", vbNullString), TitleColor(0), TitleColor(1), TitleColor(2), False, False, True
        
        If strTemp <> vbNullString Then
             AddtoRichTextBox Console, strTemp, ValueColor(0), ValueColor(1), ValueColor(2), False, False, True
        End If
        
        ' Points prize
        If .PrizePoints > 0 Then
            AddtoRichTextBox Console, "Puntos de Partida: Hasta ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox Console, .PrizePoints, ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        End If

        ' Points exp
        If .PrizeExp > 0 Then
             AddtoRichTextBox Console, "Puntos de Experiencia: Hasta ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox Console, PonerPuntos(.PrizeExp), ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        End If

        ' Level requirement
        If Not (.LvlMin = 1 And .LvlMax = 47) And Not (.LvlMin = 1 And .LvlMax = 1) Then
            'If .PrizeExp > 0 Then
                AddtoRichTextBox Console, "Nivel permitido: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
                AddtoRichTextBox Console, .LvlMin & " a " & .LvlMax & ". ", ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
            'End If
        End If

           ' Class and faction requirement
        Dim TextClass As String: TextClass = CheckAllowedClasses(.AllowedClasses)
        If TextClass <> "TODAS" Then
            AddtoRichTextBox Console, "Clases permitidas: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox Console, TextClass, ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        End If

        ' Fees
        If .InscriptionGld > 0 Or .InscriptionGldPremium > 0 Then
            AddtoRichTextBox Console, "Cuotas de inscripción: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox Console, IIf(.InscriptionGld > 0, .InscriptionGld & " de oro", "") & IIf(.InscriptionGldPremium > 0, " | " & .InscriptionGldPremium & " GldPremiums", "") & ".", ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        End If

        ' Prizes
        If .PrizeGld > 0 Or .PrizeGldPremium > 0 Or .PrizeObj.ObjIndex > 0 Then
            AddtoRichTextBox Console, "Premios: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox Console, IIf(.PrizeGld > 0, .PrizeGld & " de oro", "") & IIf(.PrizeGldPremium > 0, " | " & .PrizeGldPremium & " DSP", "") & strTemp, ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        End If

        ' Special rules
        If .Config(eConfigEvent.eCascoEscudo) = 0 Then
            AddtoRichTextBox Console, "Regla especial: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox Console, "No se permiten Cascos-Escudos.", ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        End If

        ' Special spells
        If .Config(eConfigEvent.eResu) = 1 Or .Config(eConfigEvent.eInvisibilidad) = 1 Or .Config(eConfigEvent.eOcultar) = 1 Or .Config(eConfigEvent.eInvocar) = 1 Then
            strTemp = "Hechizos NO permitidos: "
            If .Config(eConfigEvent.eResu) = 1 Then strTemp = strTemp & " 'RESU' "
            If .Config(eConfigEvent.eInvisibilidad) = 1 Then strTemp = strTemp & " 'INVI' "
            If .Config(eConfigEvent.eOcultar) = 1 Then strTemp = strTemp & " 'OCULTAR' "
            If .Config(eConfigEvent.eInvocar) = 1 Then strTemp = strTemp & " 'INVOCAR' "
            
            AddtoRichTextBox Console, strTemp, ValueColor(0), ValueColor(1), ValueColor(2), False, False, True
        End If

        ' Space
        AddtoRichTextBox Console, " ", 255, 255, 255, False, False, True
        
        Console.SelStart = 0
    End With

End Sub
