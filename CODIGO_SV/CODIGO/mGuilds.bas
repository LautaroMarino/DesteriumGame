Attribute VB_Name = "mGuilds"
' Sistema de Clanes:: EXODO II :: By Lautaro Marino (fb/lautamarino)

Option Explicit

Public Const STAT_GUILD_MAXELV       As Byte = 15

' Requisitos de talismanes

#If Classic = 0 Then
    Public Const GUILD_RUNA_VERIL        As Integer = 159
    
    Public Const GUILD_RUNA_LEVIATAN     As Integer = 160
    
    Public Const GUILD_RUNA_EGIPTO       As Integer = 161
    
    Public Const GUILD_RUNA_POLAR        As Integer = 162
    
    Public Const GUILD_RUNA_VESPAR       As Integer = 163

#Else

    Public Const GUILD_CRISTAL    As Integer = 2207

#End If
    
    Public Const MAX_GUILD_LEVEL    As Byte = 15
    
    Public Const MAX_GUILD_ALINEATION    As Byte = 5
    
    Public Const MAX_GUILD_MEMBER        As Byte = 30
    
    Public Const MIN_GUILD_POINTS        As Integer = 2000
    
    Public Const MIN_GLD_FOUND           As Long = 0
    
    Public Const MAX_GUILDS              As Integer = 300
    
    Public Const MAX_GUILD_CODEX         As Byte = 4
    
    Public Const MAX_GUILD_LEN           As Byte = 15
    
    Public Const MAX_GUILD_LEN_CODEX     As Byte = 60
    
    Public Const MAX_GUILD_RANGE         As Byte = 4
    
    Public Const MAX_GUILD_POINTS        As Integer = 10000
    
    Private Const MIN_GUILD_LEVEL_FOUND  As Byte = 35
    
    Private Const MIN_GUILD_LEVEL_MEMBER As Byte = 25

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

Public Type tGuildMember
    UserIndex As Integer
    Name As String
    Range As eGuildRange
    Char As tGuildMemberInfo

End Type

Public Type tGuild

    Name As String
    uName As String
    Date As String
    Alineation As eGuildAlineation
    Members() As tGuildMember
    LastInvitation As String
    NumMembers As Byte
    MaxMembers As Byte
    Lvl As Byte
    Elu As Long
    Exp As Long

End Type

Public GuildsInfo(0 To MAX_GUILDS) As tGuild

Public GuildLast                   As Integer

Dim FilePath_Guild                 As String

Dim FilePath_Guild_CharInfo        As String

' Fin de Declaraciones

Private Function Guild_Exist(ByVal GuildName As String) As Boolean
        '<EhHeader>
        On Error GoTo Guild_Exist_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_GUILDS

102         With GuildsInfo(A)

104             If StrComp(UCase$(.Name), GuildName) = 0 Then
106                 Guild_Exist = True
                    Exit Function

                End If

            End With

108     Next A
    
        '<EhFooter>
        Exit Function

Guild_Exist_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guild_Exist " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function Guilds_SlotUser(ByVal GuildIndex As Integer, _
                                 ByVal Name As String) As Byte

        '<EhHeader>
        On Error GoTo Guilds_SlotUser_Err

        '</EhHeader>
        Dim A As Long
    
100     With GuildsInfo(GuildIndex)

102         For A = 1 To MAX_GUILD_MEMBER

104             If StrComp(.Members(A).Name, Name) = 0 Then
106                 Guilds_SlotUser = A
                    Exit Function

                End If

108         Next A

        End With
    
        '<EhFooter>
        Exit Function

Guilds_SlotUser_Err:
        LogError Err.description & vbCrLf & "in Guilds_SlotUser " & "at line " & Erl

        '</EhFooter>
End Function

' Generamos el archivo de clanes en caso de no tenerlo.
Public Sub Create_Guilds()

    On Error GoTo ErrHandler

    Dim intFile As Integer

    Dim A       As Long, B As Long
    
    intFile = FreeFile
    
    Open FilePath_Guild For Output As #intFile
        
    Print #intFile, "[INIT]"
    Print #intFile, "GUILDLAST=0"
        
    For A = 1 To MAX_GUILDS
        Print #intFile, "[" & A & "]"
        Print #intFile, "NAME="
        Print #intFile, "ELV="
        Print #intFile, "EXP="
        Print #intFile, "ELU="
        Print #intFile, "DATE="
        Print #intFile, "ALINEATION=0"
        Print #intFile, "POINTS=0"
        Print #intFile, "MEMBERS=0"
            
        For B = 1 To MAX_GUILD_CODEX
            Print #intFile, "CODEX" & B & "="
        Next B
            
        For B = 1 To MAX_GUILD_MEMBER
            Print #intFile, "MEMBER" & B & "=-"
        Next B
    Next A
        
    Close #intFile
    
    Exit Sub

ErrHandler:
    Call LogError("ERROR Guilds:Create_Guilds: (" & Err.number & ") " & Err.description)

End Sub

' Cargamos el total de los clanes con toda su información.
Public Sub Guilds_Load()

    On Error GoTo ErrHandler

    FilePath_Guild = DatPath & "GUILDS.INI"
    
    FilePath_Guild_CharInfo = DatPath & "GUILDS_CHARINFO.INI"
    
    Dim A            As Long, B As Long

    Dim Tmp          As String

    Dim ReadGuild    As clsIniManager

    Dim ReadCharInfo As clsIniManager
    
    If Not FileExist(FilePath_Guild, vbArchive) Then
        Call Create_Guilds

    End If
    
    Set ReadGuild = New clsIniManager
    Set ReadCharInfo = New clsIniManager
    
    If FileExist(FilePath_Guild_CharInfo, vbArchive) Then
        ReadGuild.Initialize FilePath_Guild
    End If
    
    If FileExist(FilePath_Guild_CharInfo, vbArchive) Then
        ReadCharInfo.Initialize FilePath_Guild_CharInfo

    End If
        
    GuildLast = val(ReadGuild.GetValue("INIT", "GUILDLAST"))
        
    For A = 1 To MAX_GUILDS

        With GuildsInfo(A)
            .Name = ReadGuild.GetValue(A, "NAME")
            .uName = UCase$(.Name)
            .Alineation = val(ReadGuild.GetValue(A, "Alineation"))
            .Lvl = val(ReadGuild.GetValue(A, "LVL"))
            .Exp = val(ReadGuild.GetValue(A, "EXP"))
            .Elu = val(ReadGuild.GetValue(A, "ELU"))
            
            .MaxMembers = Guilds_Max_Members(.Lvl)
            
            .Date = ReadGuild.GetValue(A, "DATE")

            ReDim .Members(1 To MAX_GUILD_MEMBER) As tGuildMember
            
            .NumMembers = 0
            For B = 1 To MAX_GUILD_MEMBER
                Tmp = ReadGuild.GetValue(A, "MEMBER" & B)
                    
                .Members(B).Name = ReadField(1, Tmp, Asc("-"))
                .Members(B).Range = val(ReadField(2, Tmp, Asc("-")))
                    
                If .Members(B).Name <> vbNullString Then
                    .Members(B).Char = Guilds_Load_CharInfo(UCase$(.Members(B).Name), ReadCharInfo)
                    .NumMembers = .NumMembers + 1

                End If
                    
            Next B

        End With

    Next A
        
    Set ReadGuild = Nothing
    Set ReadCharInfo = Nothing
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes1")
    Set ReadGuild = Nothing
    Set ReadCharInfo = Nothing

End Sub

' Cargamos la información de los Chars para los renders
Public Function Guilds_Load_CharInfo(ByVal UserName As String, _
                                     ByRef Manager As clsIniManager) As tGuildMemberInfo
                                
    On Error GoTo ErrHandler

    Dim Temp As tGuildMemberInfo
    
    Temp.Name = UserName
    Temp.Body = val(Manager.GetValue(UserName, "BODY"))
    Temp.Head = val(Manager.GetValue(UserName, "HEAD"))
    Temp.Helm = val(Manager.GetValue(UserName, "HELM"))
    Temp.Shield = val(Manager.GetValue(UserName, "SHIELD"))
    Temp.Weapon = val(Manager.GetValue(UserName, "WEAPON"))
    
    Temp.Class = val(Manager.GetValue(UserName, "CLASS"))
    Temp.Raze = val(Manager.GetValue(UserName, "RAZE"))
    Temp.Elv = val(Manager.GetValue(UserName, "ELV"))
    Temp.Points = val(Manager.GetValue(UserName, "POINTS"))
    Temp.Range = val(Manager.GetValue(UserName, "RANGE"))
     
    Guilds_Load_CharInfo = Temp
    
    Exit Function

ErrHandler:
    Call LogError("Error en Clanes2")

End Function

' Guarda un clan
Public Sub Guilds_Save(ByVal GuildIndex As Integer)

    On Error GoTo ErrHandler

    Dim A         As Long, B As Long

    Dim Tmp       As String

    Dim Guild     As clsIniManager

    Dim GuildChar As clsIniManager
    
    Set Guild = New clsIniManager
    Set GuildChar = New clsIniManager
    
    Guild.Initialize FilePath_Guild
        
    If FileExist(FilePath_Guild_CharInfo, vbArchive) Then
        GuildChar.Initialize FilePath_Guild_CharInfo

    End If
        
    With GuildsInfo(GuildIndex)
        Call Guild.ChangeValue("INIT", "GUILDLAST", CStr(GuildLast))
                
        Call Guild.ChangeValue(GuildIndex, "NAME", .Name)
        Call Guild.ChangeValue(GuildIndex, "LVL", CStr(.Lvl))
        Call Guild.ChangeValue(GuildIndex, "EXP", CStr(.Exp))
        Call Guild.ChangeValue(GuildIndex, "ELU", CStr(.Elu))
        Call Guild.ChangeValue(GuildIndex, "ALINEATION", CStr(.Alineation))
        Call Guild.ChangeValue(GuildIndex, "MEMBERS", CStr(.NumMembers))
        
        For B = 1 To MAX_GUILD_MEMBER

            With .Members(B)
                Call Guild.ChangeValue(GuildIndex, "MEMBER" & B, .Name & "-" & CStr(.Range))
                        
                If .Name <> vbNullString Then
                    Call GuildChar.ChangeValue(.Name, "BODY", CStr(.Char.Body))
                    Call GuildChar.ChangeValue(.Name, "HEAD", CStr(.Char.Head))
                    Call GuildChar.ChangeValue(.Name, "HELM", CStr(.Char.Helm))
                    Call GuildChar.ChangeValue(.Name, "SHIELD", CStr(.Char.Shield))
                    Call GuildChar.ChangeValue(.Name, "WEAPON", CStr(.Char.Weapon))
                            
                    Call GuildChar.ChangeValue(.Name, "CLASS", CStr(.Char.Class))
                    Call GuildChar.ChangeValue(.Name, "RAZE", CStr(.Char.Raze))
                    Call GuildChar.ChangeValue(.Name, "ELV", CStr(.Char.Elv))
                    Call GuildChar.ChangeValue(.Name, "RANGE", CStr(.Char.Range))
                            
                End If

            End With

        Next B
    
    End With

    GuildChar.DumpFile FilePath_Guild_CharInfo
    Guild.DumpFile FilePath_Guild
        
    Set Guild = Nothing
    Set GuildChar = Nothing
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes3")
    Set Guild = Nothing
    Set GuildChar = Nothing

End Sub

Public Sub Guilds_Save_All()

    On Error GoTo ErrHandler

    Dim A          As Long, B As Long

    Dim Tmp        As String

    Dim Guild      As clsIniManager

    Dim GuildChar  As clsIniManager

    Dim GuildIndex As Integer
    
    Set Guild = New clsIniManager
    Set GuildChar = New clsIniManager
    
    
    Guild.Initialize FilePath_Guild
    GuildChar.Initialize FilePath_Guild_CharInfo
        
    For GuildIndex = 1 To MAX_GUILDS

        With GuildsInfo(GuildIndex)
            Call Guild.ChangeValue("INIT", "GUILDLAST", CStr(GuildLast))
                    
            Call Guild.ChangeValue(GuildIndex, "NAME", .Name)
            Call Guild.ChangeValue(GuildIndex, "LVL", CStr(.Lvl))
            Call Guild.ChangeValue(GuildIndex, "EXP", CStr(.Exp))
            Call Guild.ChangeValue(GuildIndex, "ELU", CStr(.Elu))
            Call Guild.ChangeValue(GuildIndex, "ALINEATION", CStr(.Alineation))
            Call Guild.ChangeValue(GuildIndex, "MEMBERS", CStr(.NumMembers))

            For B = 1 To MAX_GUILD_MEMBER

                With .Members(B)
                    Call Guild.ChangeValue(GuildIndex, "MEMBER" & B, .Name & "-" & CStr(.Range))
                            
                    If .Name <> vbNullString Then
                        Call GuildChar.ChangeValue(.Name, "BODY", CStr(.Char.Body))
                        Call GuildChar.ChangeValue(.Name, "HEAD", CStr(.Char.Head))
                        Call GuildChar.ChangeValue(.Name, "HELM", CStr(.Char.Helm))
                        Call GuildChar.ChangeValue(.Name, "SHIELD", CStr(.Char.Shield))
                        Call GuildChar.ChangeValue(.Name, "WEAPON", CStr(.Char.Weapon))
                                
                        Call GuildChar.ChangeValue(.Name, "CLASS", CStr(.Char.Class))
                        Call GuildChar.ChangeValue(.Name, "RAZE", CStr(.Char.Raze))
                        Call GuildChar.ChangeValue(.Name, "ELV", CStr(.Char.Elv))
                        Call GuildChar.ChangeValue(.Name, "RANGE", CStr(.Char.Range))
                                
                    End If

                End With

            Next B

        End With

    Next GuildIndex
        
    GuildChar.DumpFile FilePath_Guild_CharInfo
    Guild.DumpFile FilePath_Guild
        
    Set Guild = Nothing
    Set GuildChar = Nothing
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes4")

End Sub

' Realiza las comprobaciones necesarias para la fundación del clan.
Private Function Guilds_Check_New(ByVal UserIndex As Integer, _
                                  ByVal Name As String, _
                                  ByVal Alineation As eGuildAlineation, _
                                  ByRef Codex() As String) As String

    On Error GoTo ErrHandler

    Dim A As Long
    
    With UserList(UserIndex)

        If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
            Guilds_Check_New = "El nombre de tu clan contiene caracteres inválidos."
            Exit Function

        End If
    
        If Right$(Name, 1) = " " Then
            Guilds_Check_New = "Tu nivel no te permite fundar un clan."
            
            Exit Function

        End If
        
        If .Stats.Elv < MIN_GUILD_LEVEL_FOUND Then
            Guilds_Check_New = "Necesitas ser nivel " & MIN_GUILD_LEVEL_FOUND & " para fundar clan."

            Exit Function

        End If
        
        If .GuildIndex > 0 Then
            Guilds_Check_New = "Ya perteneces a un clan."

            Exit Function

        End If
        
        If GuildLast = MAX_GUILDS Then
            Guilds_Check_New = "No hay lugar para nuevos Clanes. ¡Han sido todos creados!"

            Exit Function

        End If
        
        If MapInfo(.Pos.Map).Pk Then
            Guilds_Check_New = "Para fundar un clan debes dirigirte a zona segura"

            Exit Function

        End If
        
        If Len(Name) <= 0 Or Len(Name) > MAX_GUILD_LEN Then
            Guilds_Check_New = "Para fundar un clan debes dirigirte a zona segura"
            'Call Logs_Security(eSecurity, eAntiHack, "Guilds_Check_New::NAME:: El personaje " & UserList(UserIndex).Name & " con IP: " & .Ip & " Y HD: 0  intentó hackear el sistema")

            Exit Function

        End If
        
        'For A = LBound(Codex) To UBound(Codex)

        ' If Len(Codex(A)) <= 0 Or Len(Codex(A)) > MAX_GUILD_LEN_CODEX Then
        'Guilds_Check_New = "Para fundar un clan debes dirigirte a zona segura"
        'Call Logs_Security(eSecurity, eAntiHack, "Guilds_Check_New::CODEX:: El personaje " & UserList(UserIndex).Name & " con IP: " & .Ip & " Y HD: 0  intentó hackear el sistema")

        ' Exit Function

        'End If

        'Next A
        
        If Alineation < 0 Or Alineation > MAX_GUILD_ALINEATION Then
            Guilds_Check_New = "Para fundar un clan debes dirigirte a zona segura"
            'Call Logs_Security(eSecurity, eAntiHack, "Guilds_Check_New::ALINEATION:: El personaje " & UserList(UserIndex).Name & " con IP: " & .Ip & " Y HD: 0  intentó hackear el sistema")

            Exit Function

        End If
        
        If Not Guilds_Check_Alineation(UserIndex, Alineation) Then
            Guilds_Check_New = "Tu personaje posee una alineación distinta a la que has elegido."

            Exit Function

        End If
        
            If Not TieneObjetos(GUILD_CRISTAL, 3, UserIndex) Then
                Guilds_Check_New = "Consigue 3 Cristales de Hielo. ¡Se crean con 10 Fragmentos de Cristal!"

                Exit Function

            End If
        
            If .Stats.Gld < MIN_GLD_FOUND Then
                Guilds_Check_New = "Para fundar un clan debes disponer de " & MIN_GLD_FOUND & " Monedas de Oro."

                Exit Function

            End If
       
            If Guild_Exist(UCase$(Name)) Then
                Guilds_Check_New = "Ya existe un clan con ese nombre."
                Exit Function

            End If

        End With
    
        Exit Function

ErrHandler:
        Call LogError("Error en Clanes5")
    
    End Function
    
' # Determina el máximo de miembros permitidos por Nivel
Function Guilds_Max_Members(ByVal Elv As Integer) As Integer
    Select Case Elv
        Case 1
            Guilds_Max_Members = 6
        Case 2, 3
            Guilds_Max_Members = 9
        Case 4
            Guilds_Max_Members = 11
        Case 5
            Guilds_Max_Members = 13
        Case 6 To 7
            Guilds_Max_Members = 15
        Case 8
            Guilds_Max_Members = 17
        Case 9
            Guilds_Max_Members = 19
        Case 10
            Guilds_Max_Members = 21
        Case 11 To 12
            Guilds_Max_Members = 24
        Case 13
            Guilds_Max_Members = 27
        Case 14 To 15
            Guilds_Max_Members = 30
        Case Else
            Guilds_Max_Members = 0 ' Nivel no válido
    End Select
End Function
' Se funda un nuevo clan
Public Sub Guilds_New(ByVal UserIndex As Integer, _
                      ByVal Name As String, _
                      ByVal Alineation As eGuildAlineation, _
                      ByRef Codex() As String)

    On Error GoTo ErrHandler

    Dim ErrorMsg As String: ErrorMsg = Guilds_Check_New(UserIndex, Name, Alineation, Codex)
    
    If ErrorMsg <> vbNullString Then
        Call WriteConsoleMsg(UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING)

        Exit Sub

    End If
    
    GuildLast = GuildLast + 1

    With GuildsInfo(GuildLast)
        .uName = UCase$(Name)
        .Name = Name
        .Alineation = Alineation
        .Lvl = 1
        .Exp = 0
        .Elu = Guilds_Elu(.Lvl)
        
        .MaxMembers = Guilds_Max_Members(.Lvl)
        
        ReDim .Members(1 To MAX_GUILD_MEMBER) As tGuildMember
        
        Call Guilds_Add_Member(GuildLast, UserIndex, rFound)
        
    End With
    
    Call QuitarObjetos(GUILD_CRISTAL, 3, UserIndex)
    
    UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - MIN_GLD_FOUND
    Call WriteUpdateGold(UserIndex)
    'Call mRank.RankUser_AddPoint(UserIndex, 250)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje " & UserList(UserIndex).Name & " ha fundado el clan '" & Name & "' de alineación '" & Guilds_Alineation_String(Alineation) & "'", FontTypeNames.FONTTYPE_GUILD))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(44, NO_3D_SOUND, NO_3D_SOUND))
    Call Guilds_Save(GuildLast)
    Call RefreshCharStatus(UserIndex)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes6")

End Sub

' # Determina la experiencia a juntar para pasar al siguiente Nivel
Private Function Guilds_Elu(ByVal Lvl As Byte) As Long
    Dim experienceNeeded As Long
    
    Select Case Lvl
        Case 1
            experienceNeeded = 50
        Case Else
            experienceNeeded = 100 * (Lvl - 1)
    End Select
    
    Guilds_Elu = experienceNeeded
End Function


' Enviamos una invitación a cualquier personaje para que ingrese a nuestro clan.
Public Sub Guilds_SendInvitation(ByVal UserIndex As Integer, ByVal UserName As String)

    On Error GoTo ErrHandler

    Dim tUser      As String

    Dim GuildIndex As Integer
    
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Then Exit Sub
    
    With GuildsInfo(GuildIndex)

        If UserList(UserIndex).GuildRange <> rFound And UserList(UserIndex).GuildRange <> rLeader Then
            Call WriteConsoleMsg(UserIndex, "Solo el Líder del clan puede invitar nuevos miembros", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If .NumMembers = MAX_GUILD_MEMBER Then
            Call WriteConsoleMsg(UserIndex, "El clan alcanzó el límite de miembros", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If Not PersonajeExiste(UserName) Then
            Call WriteConsoleMsg(UserIndex, "El usuario no existe.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        tUser = NameIndex(UserName)
        
        Dim Lvl As Byte
        Dim GuildUser As Byte
        
        If tUser <= 0 Then
            Lvl = val(GetVar(CharPath & UCase$(UserName) & ".chr", "STATS", "ELV"))
            GuildUser = val(GetVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GUILDINDEX"))
        Else
            GuildUser = UserList(tUser).GuildIndex
            Lvl = UserList(tUser).Stats.Elv


        End If
        
        If Lvl < MIN_GUILD_LEVEL_MEMBER Then
            Call WriteConsoleMsg(UserIndex, "El usuario no tiene suficiente nivel para participar de un clan.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If GuildUser > 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario ya pertenece a un clan.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        .LastInvitation = UCase$(UserName)
        
        Call WriteConsoleMsg(UserIndex, "Le has enviado una solicitud al usuario " & UserName & ". Recuerda que si envias solicitud a alguien más, la última enviada se borrará.", FontTypeNames.FONTTYPE_INFOGREEN)
        
        If tUser > 0 Then
            Call WriteConsoleMsg(tUser, "El lider del clan " & .Name & " te ha ofrecido pertenecer a su clan. Si deseas aceptar la invitación tipea /SICLAN " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_INFOGREEN)
        End If
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes7")

End Sub

' El personaje al que ofrecimos unirse a nuestro clan, nos acepta.
Public Sub Guilds_AcceptInvitation(ByVal UserIndex As Integer, ByVal LeaderName As String)

    On Error GoTo ErrHandler

    Dim GuildIndex As Integer

    Dim GuildName  As String
    
    Dim GuildRange As Byte

    Dim tLeader    As Integer
    
    Dim Slot       As Byte

    If UserList(UserIndex).GuildIndex > 0 Then
        Call WriteConsoleMsg(UserIndex, "Ya perteneces a un clan.", FontTypeNames.FONTTYPE_INFORED)

        Exit Sub

    End If
    
    tLeader = NameIndex(LeaderName)
    
    If tLeader <= 0 Then
        GuildIndex = val(GetVar(CharPath & UCase$(LeaderName) & ".chr", "GUILD", "GUILDINDEX"))
        GuildRange = val(GetVar(CharPath & UCase$(LeaderName) & ".chr", "GUILD", "GUILDRANGE"))
    Else
        GuildIndex = UserList(tLeader).GuildIndex
        GuildRange = UserList(tLeader).GuildRange
    End If
    
    If GuildIndex = 0 Then
        Call WriteConsoleMsg(UserIndex, "El personaje no pertenece a ningún clan.", FontTypeNames.FONTTYPE_INFORED)

        Exit Sub

    End If
    
    With GuildsInfo(GuildIndex)

        If .LastInvitation <> UCase$(UserList(UserIndex).Name) Then
            Call WriteConsoleMsg(UserIndex, "El Líder no te ha invitado a su clan.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If GuildRange <> rFound And GuildRange <> rLeader Then
            Call WriteConsoleMsg(UserIndex, "El Líder no te ha invitado a su clan.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If .NumMembers = .MaxMembers Then
            Call WriteConsoleMsg(UserIndex, "El clan alcanzó el límite de miembros según tu nivel.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
    
        If Not Guilds_Check_Alineation(UserIndex, .Alineation) Then
            Call WriteConsoleMsg(UserIndex, "Tu alineación no te permite entrar al clan.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        .LastInvitation = vbNullString
        
        Call Guilds_Add_Member(GuildIndex, UserIndex, rNone)
        Call RefreshCharStatus(UserIndex)
        Call SendData(SendTarget.ToDiosesYclan, GuildIndex, PrepareMessagePlayEffect(43, NO_3D_SOUND, NO_3D_SOUND))
        Call SendData(SendTarget.ToDiosesYclan, GuildIndex, PrepareMessageConsoleMsg("Damos la bienvenida al personaje " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_GUILDMSG))

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes8")

End Sub

' Un personaje abandona un clan
Public Sub Guilds_KickMe(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Guilds_KickMe_Err
        '</EhHeader>

        On Error GoTo ErrHandler
    
        Dim GuildIndex As Integer
    
100     With UserList(UserIndex)
102         GuildIndex = .GuildIndex
            
104         If .GuildRange = rFound Or .GuildRange = rLeader Then
106             Call WriteConsoleMsg(UserIndex, "¡Tú no puedes salir del clan!", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If
            
             
             If .GuildSlot > 0 Then
108             Call Guilds_KickUserSlot(GuildIndex, .GuildSlot)
             Else
                   .GuildIndex = 0
                   .GuildSlot = 0
                   .GuildRange = 0
             End If
        End With
    
        Exit Sub

ErrHandler:
110     Call LogError("Error en Clanes9")

        '<EhFooter>
        Exit Sub

Guilds_KickMe_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_KickMe " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' El Líder/Fundador elimina a un personaje del clan
Public Sub Guilds_KickUser(ByVal UserIndex As Integer, ByVal UserName As String)

    On Error GoTo ErrHandler

    ' // NUEVO
    
    Dim tUser As Integer

    Dim SlotUser As Integer

    Dim TempGuild As Integer
    
    With UserList(UserIndex)

        Dim GuildIndex As Integer

        GuildIndex = .GuildIndex
        
        If GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No perteneces a ningún clan!", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If Not .GuildRange = rFound And Not .GuildRange = rLeader Then
            Call WriteConsoleMsg(UserIndex, "¡Solo el Líder puede realizar esta acción!", FontTypeNames.FONTTYPE_INFORED)
            
            Exit Sub

        End If
        
        If UCase$(UserName) = UCase$(.Name) Then
            Call WriteConsoleMsg(UserIndex, "¡No puedes echarte a ti mismo!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
        
        tUser = NameIndex(UserName)
        
        If tUser > 0 Then
            TempGuild = UserList(tUser).GuildIndex
        Else
            TempGuild = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
        End If

        SlotUser = Guilds_Search_User(GuildIndex, UCase$(UserName))
        
        If SlotUser > 0 Then
            
            ' # FIX :: EL PERSONAJE NO PERTENECE AL CLAN PERO SIGUE FIGURETI
            If TempGuild <> GuildIndex Then
                Dim Nulo As tGuildMember
                GuildsInfo(GuildIndex).Members(SlotUser) = Nulo
                GuildsInfo(GuildIndex).NumMembers = GuildsInfo(GuildIndex).NumMembers - 1
                Call Guilds_Save(GuildIndex)
                Call LogError("Clanes:: Clan " & GuildIndex & " Encontró un personaje bug " & UserName)
                Exit Sub
            End If
            
            Call Guilds_KickUserSlot(GuildIndex, SlotUser)
            Call WriteConsoleMsg(UserIndex, "¡Has expulsado a " & UserName & "!", FontTypeNames.FONTTYPE_INFOGREEN)
        End If
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes10")

End Sub

' Chequeos y preparación de información para el LIDER.
Public Sub Guilds_PrepareInfoUsers(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    Dim A        As Long

    Dim FilePath As String
    
    With UserList(UserIndex)

        If .GuildIndex = 0 Then
            Call Logs_Security(eSecurity, eAntiHack, "Guilds_PrepareInfoUsers::CLAN NULO:: El personaje " & .Name & " con IP: " & .IpAddress & " ha solicitado el panel teniendo clan inválido")

            Exit Sub

        End If
        
        If .GuildRange <> rFound And .GuildRange <> rLeader Then
            Call Logs_Security(eSecurity, eAntiHack, "Guilds_PrepareInfoUsers::NO LIDER:: El personaje " & .Name & " con IP: " & .IpAddress & " ha solicitado el panel no siendo un personaje Líder")

            Exit Sub

        End If
        
        Call WriteGuild_InfoUsers(UserIndex, .GuildIndex, GuildsInfo(.GuildIndex).Members)

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes11")

End Sub

' Buscamos un rango especial (LIDER, VOCERO) y sacamos al anterior.
Private Sub Guilds_RemoveRange(ByVal GuildIndex As Integer, _
                               ByVal SlotUser As Byte, _
                               ByVal Range As eGuildRange)

    On Error GoTo ErrHandler

    Dim A     As Long

    Dim tUser As Integer

    For A = 1 To MAX_GUILD_MEMBER

        With GuildsInfo(GuildIndex).Members(A)

            If .Name <> vbNullString Then
                If .Range <> rNone And .Range = Range Then
                    tUser = NameIndex(.Name)
                    
                    If tUser > 0 Then
                        Call Guilds_Update_Member_Range(tUser, 0)
                    Else
                        Call Guilds_Save_Member_Range(UCase$(.Name), 0)

                    End If
                    
                    Exit For

                End If

            End If

        End With

    Next A
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes13")

End Sub

' Cuando hay un cambio de alineación, al ingresar a un personaje que pertenece a ese clan, chequeamos si es válido de permanecer en él o no.
Public Sub Guilds_CheckAlineation(ByVal UserIndex As Integer, _
                                  Optional ByVal NewAlineation As eGuildAlineation = eGuildAlineation.a_Neutral)

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        Dim Slot As Byte
        
        If .GuildIndex = 0 Then Exit Sub
        
        Slot = Guilds_SearchSlotUser(.GuildIndex, UCase$(.Name))
        
        If Not Guilds_Check_Alineation(UserIndex, GuildsInfo(.GuildIndex).Alineation) Then
            If .GuildRange = eGuildRange.rFound Then
                Call Guilds_ChangeAlineation(.GuildIndex, NewAlineation)
            
            Else
                Call WriteConsoleMsg(UserIndex, "Tu alineación no te permite permanecer en el clan.", FontTypeNames.FONTTYPE_INFORED)
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("El personaje " & .Name & " no ha podido validar su permanencia en el clan.", FontTypeNames.FONTTYPE_GUILDMSG))
                            
                 If Slot > 0 Then
                    Call Guilds_KickUserSlot(.GuildIndex, Slot)
                End If
                
            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes14")

End Sub

' Cambiamos la alineación de un clan (Debido al Lider)
Public Sub Guilds_ChangeAlineation(ByVal GuildIndex As Integer, _
                                   ByVal Alineation As eGuildAlineation)

    On Error GoTo ErrHandler

    With GuildsInfo(GuildIndex)
        .Alineation = Alineation
        
        Call SendData(SendTarget.ToDiosesYclan, GuildIndex, PrepareMessageConsoleMsg("La alineación del clan ha pasado a ser '" & Guilds_Alineation_String(Alineation) & "'", FontTypeNames.FONTTYPE_GUILDMSG))
        
        Call Guilds_ValidatePermanencia(GuildIndex, Alineation)

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes15")

End Sub

' Valida la permanencia de los personajes online en el clan debido a su cambio de alineación.
Private Sub Guilds_ValidatePermanencia(ByVal GuildIndex As Integer, _
                                       ByVal Alineation As eGuildAlineation)

    Dim A     As Long

    Dim tUser As Integer

    On Error GoTo ErrHandler

    With GuildsInfo(GuildIndex)

        For A = 2 To MAX_GUILD_MEMBER

            With .Members(A)
                tUser = NameIndex(.Name)
                
                If tUser > 0 Then
                    If Not Guilds_Check_Alineation(tUser, Alineation) Then
                        Call WriteConsoleMsg(tUser, "Tu alineación no te permite permanecer en el clan.", FontTypeNames.FONTTYPE_INFORED)
                        Call SendData(SendTarget.ToDiosesYclan, GuildIndex, PrepareMessageConsoleMsg("El personaje " & .Name & " no ha podido validar su permanencia en el clan.", FontTypeNames.FONTTYPE_GUILDMSG))
                        Call Guilds_KickUserSlot(GuildIndex, A)

                    End If

                End If
                
            End With
            
        Next A
    
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes16")

End Sub

' #############
' ############# Funciones y procedimientos de los clanes que se utilizan sin modificaciones (En la mayoría de los casos)
' #############

' Función que devuelve el texto de la alineación, en base a su número identificador.
Public Function Guilds_Alineation_String(ByVal Alineation As eGuildAlineation) As String

    Select Case Alineation

        Case eGuildAlineation.a_Armada: Guilds_Alineation_String = "Real"

        Case eGuildAlineation.a_Legion: Guilds_Alineation_String = "Caos"

        Case eGuildAlineation.a_Neutral: Guilds_Alineation_String = "Neutral"

    End Select

End Function

' Valida TRUE o FALSE según la alineación del clan elegido y la alineación del personaje)
Private Function Guilds_Check_Alineation(ByVal UserIndex As Integer, _
                                         ByVal Alineation As eGuildAlineation) As Boolean
                                         
    Select Case Alineation
    
        Case eGuildAlineation.a_Armada
            Guilds_Check_Alineation = (UserList(UserIndex).Faction.Status = r_Armada)
            
        Case eGuildAlineation.a_Legion
            Guilds_Check_Alineation = (UserList(UserIndex).Faction.Status = r_Caos)
        
        Case eGuildAlineation.a_Neutral
            Guilds_Check_Alineation = (UserList(UserIndex).Faction.Status <> r_Caos And UserList(UserIndex).Faction.Status <> r_Armada)
            
    End Select

End Function

' Preparamos el nombre del rango elegido
Public Function Guilds_PrepareRangeName(ByVal Range As eGuildRange) As String

    Select Case Range
    
        Case eGuildRange.rNone
            Guilds_PrepareRangeName = "Miembro"
            
        Case eGuildRange.rFound
            Guilds_PrepareRangeName = "Fundador"
            
        Case eGuildRange.rLeader
            Guilds_PrepareRangeName = "Lider"
            
        Case eGuildRange.rVocero
            Guilds_PrepareRangeName = "Vocero"

    End Select

End Function

' Preparamos la lista de todos los clanes disponible.
Public Function Guilds_PrepareList() As String()
        '<EhHeader>
        On Error GoTo Guilds_PrepareList_Err
        '</EhHeader>

        Dim A                          As Long

        Dim GuildList(1 To MAX_GUILDS) As String
        
100     For A = 1 To MAX_GUILDS
102         GuildList(A) = GuildsInfo(A).Name
104     Next A
    
106     Guilds_PrepareList = GuildList

        '<EhFooter>
        Exit Function

Guilds_PrepareList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_PrepareList " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Buscamos un personaje en una GUILD.
Private Function Guilds_SearchSlotUser(ByVal GuildIndex As Integer, _
                                       ByVal UserName As String) As Byte
        '<EhHeader>
        On Error GoTo Guilds_SearchSlotUser_Err
        '</EhHeader>
    
        Dim A As Long
    
100     For A = 1 To MAX_GUILD_MEMBER

102         With GuildsInfo(GuildIndex).Members(A)

104             If .Name = UserName Then
106                 Guilds_SearchSlotUser = A

                    Exit Function

                End If

            End With

108     Next A

        '<EhFooter>
        Exit Function

Guilds_SearchSlotUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_SearchSlotUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Preparamos la lista de usuarios ONLINE/OFFLINE
Public Sub Guilds_PrepareOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)

    On Error GoTo ErrHandler

    Dim A           As Long

    Dim TempOnline  As String

    Dim TempOffline As String

    For A = 1 To MAX_GUILD_MEMBER

        With GuildsInfo(GuildIndex).Members(A)

            If .Name <> vbNullString Then
                If CheckUserLogged(UCase$(.Name)) Then
                    TempOnline = TempOnline & .Name & ", "
                Else
                    TempOffline = TempOffline & .Name & ", "

                End If

            End If

        End With

    Next A
    
    If Len(TempOnline) > 0 Then TempOnline = Left$(TempOnline, Len(TempOnline) - 2)
    
    If Len(TempOffline) > 0 Then TempOffline = Left$(TempOffline, Len(TempOffline) - 2)
    
    Call WriteConsoleMsg(UserIndex, "Usuarios conectados: " & TempOnline, FontTypeNames.FONTTYPE_INFOGREEN)
    Call WriteConsoleMsg(UserIndex, "Usuarios offline: " & TempOffline, FontTypeNames.FONTTYPE_INFORED)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes19")

End Sub

' Eliminamos a un personaje de la GUILD
Private Sub Guilds_KickUserSlot(ByVal GuildIndex As Integer, ByVal SlotUser As Byte)

        '<EhHeader>
        On Error GoTo Guilds_KickUserSlot_Err

        '</EhHeader>
        Dim UserName                    As String


        Dim A                           As Long, B As Long

        Dim Temp(1 To MAX_GUILD_MEMBER) As tGuildMember

        Dim TempVacio                   As tGuildMember
        
        Dim tUser As Integer
        
100     With GuildsInfo(GuildIndex)
102         UserName = .Members(SlotUser).Name
        
104         tUser = NameIndex(UserName)
        
106         If tUser > 0 Then
108             Call Guilds_Update_Member(tUser, 0)
110             Call Guilds_Update_Member_Range(tUser, 0)
                  
                  GuildsInfo(GuildIndex).Members(SlotUser).UserIndex = 0
                  UserList(tUser).GuildSlot = 0
112             Call RefreshCharStatus(tUser)
            Else
114             Call Guilds_Save_Member(UCase$(.Members(SlotUser).Name), 0)
116             Call Guilds_Save_Member_Range(UCase$(.Members(SlotUser).Name), 0)

            End If

124         With .Members(SlotUser)
126             .Name = vbNullString
128             .Range = rNone

130             .Char.Body = 0
132             .Char.Head = 0
134             .Char.Weapon = 0
136             .Char.Helm = 0
138             .Char.Class = 0
140             .Char.Raze = 0
142             .Char.Elv = 0
144             .Char.Range = 0
146             .Char.Name = vbNullString
148             .Char.Points = 0

            End With
        
158         .NumMembers = .NumMembers - 1
            
              Call mGuilds.Guilds_Save(GuildIndex)
        End With

        '<EhFooter>
        Exit Sub

Guilds_KickUserSlot_Err:
        LogError Err.description & vbCrLf & "in Guilds_KickUserSlot " & "at line " & Erl

        '</EhFooter>
End Sub

Private Function Guilds_Search_User_FreeSlot(ByVal GuildIndex As Integer) As Byte
        '<EhHeader>
        On Error GoTo Guilds_Search_User_FreeSlot_Err
        '</EhHeader>

        Dim A As Long
    
100     With GuildsInfo(GuildIndex)

102         For A = 1 To MAX_GUILD_MEMBER

104             If .Members(A).Name = vbNullString Then
106                 Guilds_Search_User_FreeSlot = A
                    Exit Function

                End If

108         Next A
    
        End With

        '<EhFooter>
        Exit Function

Guilds_Search_User_FreeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Search_User_FreeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function Guilds_Search_User(ByVal GuildIndex As Integer, ByVal UserName As String) As Byte
        '<EhHeader>
        On Error GoTo Guilds_Search_User_Err
        '</EhHeader>

        Dim A As Long
    
100     With GuildsInfo(GuildIndex)

102         For A = 1 To MAX_GUILD_MEMBER

104             If StrComp(.Members(A).Name, UserName) = 0 Then
106                 Guilds_Search_User = A
                    Exit Function

                End If

108         Next A
    
        End With

        '<EhFooter>
        Exit Function

Guilds_Search_User_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Search_User " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
' Se agrega un personaje a un clan
Private Sub Guilds_Add_Member(ByVal GuildIndex As Integer, _
                              ByVal UserIndex As Integer, _
                              ByVal GuildRange As eGuildRange)

    On Error GoTo ErrHandler

    Dim Slot As Byte

    With GuildsInfo(GuildIndex)
        Slot = Guilds_Search_User_FreeSlot(GuildIndex)
        
        .NumMembers = .NumMembers + 1
        
        Call Guilds_Modify_Member(GuildIndex, Slot, UserList(UserIndex).Name, GuildRange)
        Call Guilds_Update_Member(UserIndex, GuildIndex)
        Call Guilds_Update_Member_Range(UserIndex, GuildRange)
        Call Guilds_Update_Char_Info(UserIndex)
        
        Call SendData(SendTarget.ToOne, UserIndex, PrepareUpdateLevelGuild(.Lvl))
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes21")

End Sub

' Modificamos un miembro del clan en base a su SLOT.
Private Sub Guilds_Modify_Member(ByVal GuildIndex As Integer, _
                                 ByVal SlotMember As Byte, _
                                 ByVal UserName As String, _
                                 ByVal GuildRange As eGuildRange)
        '<EhHeader>
        On Error GoTo Guilds_Modify_Member_Err
        '</EhHeader>
                                 
100     With GuildsInfo(GuildIndex)
102         .Members(SlotMember).Name = UCase$(UserName)
104         .Members(SlotMember).Range = GuildRange

        End With

        '<EhFooter>
        Exit Sub

Guilds_Modify_Member_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Modify_Member " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Actualizamos la información del personaje. (GUILDINDEX)
Private Sub Guilds_Update_Member(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
        '<EhHeader>
        On Error GoTo Guilds_Update_Member_Err
        '</EhHeader>
                                 
100     With UserList(UserIndex)
102         .GuildIndex = GuildIndex
104         Call Guilds_Save_Member(UCase$(.Name), GuildIndex)

        End With

        '<EhFooter>
        Exit Sub

Guilds_Update_Member_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Update_Member " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Guardamos la información del personaje (GUILDINDEX)
Private Sub Guilds_Save_Member(ByVal UserName As String, ByVal GuildIndex As Integer)
        '<EhHeader>
        On Error GoTo Guilds_Save_Member_Err
        '</EhHeader>

100     Call WriteVar(CharPath & UserName & ".chr", "GUILD", "GuildIndex", CStr(GuildIndex))

        '<EhFooter>
        Exit Sub

Guilds_Save_Member_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Save_Member " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Actualizamos la información del personaje. (RANGO)
Private Sub Guilds_Update_Member_Range(ByVal UserIndex As Integer, _
                                       ByVal GuildRange As eGuildRange)
        '<EhHeader>
        On Error GoTo Guilds_Update_Member_Range_Err
        '</EhHeader>
                                 
100     With UserList(UserIndex)
102         .GuildRange = GuildRange
        
104         Call Guilds_Save_Member_Range(UCase$(.Name), GuildRange)

        End With

        '<EhFooter>
        Exit Sub

Guilds_Update_Member_Range_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Update_Member_Range " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Guardamos la información del personaje. (RANGO)
Private Sub Guilds_Save_Member_Range(ByVal UserName As String, _
                                     ByVal GuildRange As eGuildRange)
        '<EhHeader>
        On Error GoTo Guilds_Save_Member_Range_Err
        '</EhHeader>
                                 
100     Call WriteVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GuildRange", CStr(GuildRange))

        '<EhFooter>
        Exit Sub

Guilds_Save_Member_Range_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_Save_Member_Range " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Chequeo inicial. Esto se debe a que los clanes se actualizan cada worldsave y pueden haber rollbacks
Private Function Guilds_CheckRollBack(ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Guilds_CheckRollBack_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

            Dim SlotUser As Byte

102         SlotUser = Guilds_SearchSlotUser(.GuildIndex, UCase$(.Name))
        
104         If SlotUser = 0 Then
106             Call Guilds_Update_Member(UserIndex, 0)
108             Guilds_CheckRollBack = True

            End If
        
        End With

        '<EhFooter>
        Exit Function

Guilds_CheckRollBack_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_CheckRollBack " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Un miembro de un clan se conecta
Public Sub Guilds_Connect(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If Guilds_CheckRollBack(UserIndex) Then Exit Sub

        Call Guilds_CheckAlineation(UserIndex)
        
        ' ¿Despues de chequear la alineación sigue teniendo clan?
        If .GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("El personaje " & .Name & " se ha conectado.", FontTypeNames.FONTTYPE_GUILDMSG))
            Call Guilds_Update_Char_Info(UserIndex)
            Call SendData(SendTarget.ToOne, UserIndex, PrepareUpdateLevelGuild(GuildsInfo(.GuildIndex).Lvl))
        End If

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes22")
    
End Sub

' Actualizamos la información para los renders
Private Sub Guilds_Update_Char_Info(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    Dim Slot As Byte

    Slot = Guilds_SearchSlotUser(UserList(UserIndex).GuildIndex, UCase$(UserList(UserIndex).Name))
    
    ' Como guardamos los clanes cada 1 hora, aveces se puede generar un rollback, es por eso que al logear chequeamos la situación del personaje.
    If Slot = 0 Then
        Call Guilds_Update_Member(UserIndex, 0)
        Call Guilds_Update_Member_Range(UserIndex, rNone)
        
        Call WriteConsoleMsg(UserIndex, "No hemos podido comprobar tu permanencia en el clan. ¡Has sido expulsado!", FontTypeNames.FONTTYPE_INFORED)
    Else
        
        ' Actualizo el SLOT para usar Integers en vez de Strings
        UserList(UserIndex).GuildSlot = Slot
        GuildsInfo(UserList(UserIndex).GuildIndex).Members(Slot).UserIndex = UserIndex
        
        With GuildsInfo(UserList(UserIndex).GuildIndex).Members(Slot).Char
            
            .Name = UserList(UserIndex).Name
            
            'If UserList(UserIndex).flags.Mimetizado = 1 Then
            '.Body = UserList(UserIndex).CharMimetizado.Body
            '.Head = UserList(UserIndex).CharMimetizado.Head
            '.Helm = UserList(UserIndex).CharMimetizado.CascoAnim
            '.Shield = UserList(UserIndex).CharMimetizado.ShieldAnim
            '.Weapon = UserList(UserIndex).CharMimetizado.WeaponAnim
            'Else
            .Body = UserList(UserIndex).OrigChar.Body
            .Head = UserList(UserIndex).OrigChar.Head
            .Helm = UserList(UserIndex).OrigChar.CascoAnim
            .Shield = UserList(UserIndex).OrigChar.ShieldAnim
            .Weapon = UserList(UserIndex).OrigChar.WeaponAnim
            
            'End If
            
            .Class = UserList(UserIndex).Clase
            .Raze = UserList(UserIndex).Raza
            .Elv = UserList(UserIndex).Stats.Elv
            .Points = UserList(UserIndex).Stats.Points
            .Range = UserList(UserIndex).GuildRange

        End With

    End If
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Clanes23")

End Sub

Public Function Guilds_SearchIndex(ByVal GuildName As String) As Integer
        '<EhHeader>
        On Error GoTo Guilds_SearchIndex_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_GUILDS

102         With GuildsInfo(A)

104             If StrComp(UCase$(GuildsInfo(A).Name), GuildName) = 0 Then
106                 Guilds_SearchIndex = A
                    Exit Function

                End If
        
            End With
    
108     Next A
    
        '<EhFooter>
        Exit Function

Guilds_SearchIndex_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_SearchIndex " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' El clan suma experiencia y/o nivel
Public Sub Guilds_AddExp(ByVal UserIndex As Integer, ByVal Exp As Long)
        '<EhHeader>
        On Error GoTo Guilds_AddExp_Err
        '</EhHeader>

        Dim GuildIndex As Integer

100     GuildIndex = UserList(UserIndex).GuildIndex
    
102     If GuildsInfo(GuildIndex).Lvl = STAT_GUILD_MAXELV Then Exit Sub
    
104     With GuildsInfo(GuildIndex)
         
106         .Exp = .Exp + Exp
        
108         Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageRenderConsole("ExpClan +" & CStr(Format(Exp, "###,###,###")), d_Exp, 3000, 0))
        
110         Do While .Exp >= .Elu
            
                'Checkea si alcanzó el máximo nivel
112             If .Lvl >= STAT_GUILD_MAXELV Then
114                 .Exp = 0
116                 .Elu = 0
                    Exit Sub

                End If
    
118             .Lvl = .Lvl + 1
120             .Exp = .Exp - .Elu
122             .Elu = Guilds_Elu(.Lvl)
124             .MaxMembers = Guilds_Max_Members(.Lvl)
            
126             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Clan " & .Name & " ha llegado al Nivel '" & .Lvl & "'", FontTypeNames.FONTTYPE_GUILD))
128             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(eSound.sVictory4, 50, 50, UserList(UserIndex).Char.charindex))
130             Call SendData(SendTarget.ToDiosesYclan, GuildIndex, PrepareUpdateLevelGuild(.Lvl))
132             Call Guilds_Save(GuildIndex)
            Loop
   
        End With

        '<EhFooter>
        Exit Sub

Guilds_AddExp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mGuilds.Guilds_AddExp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Enviamos la Posicion del UserIndex a los Miembros Activos
Public Sub Guilds_UpdatePosition(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo Guilds_UpdatePosition_Err

        '</EhHeader>

        Dim A          As Long

        Dim GuildIndex As Integer
    
100     GuildIndex = UserList(UserIndex).GuildIndex
    
102     If GuildIndex = 0 Then Exit Sub
    
104     With GuildsInfo(GuildIndex)
    
            ' Minimo de Nivel Requerido >>
106         If GuildsInfo(GuildIndex).Lvl < 7 Then Exit Sub
        
108         For A = 1 To MAX_GUILD_MEMBER

110             If .Members(A).UserIndex > 0 And .Members(A).UserIndex <> UserIndex Then
112                   If UserList(.Members(A).UserIndex).Pos.Map = UserList(UserIndex).Pos.Map Then
114                         Call WriteUpdatePosGuild(.Members(A).UserIndex, A, UserIndex)
                            
                        Else
116                         Call WriteUpdatePosGuild(.Members(A).UserIndex, A, 0)
        
                        End If

                End If

118         Next A
    
        End With
    
        '<EhFooter>
        Exit Sub

Guilds_UpdatePosition_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mGuilds.Guilds_UpdatePosition " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub ChangeLeader(ByVal UserIndex As Integer, ByVal UserName As String)

        '<EhHeader>
        On Error GoTo ChangeLeader_Err

        '</EhHeader>
    
        Dim tUser     As Integer

        Dim SlotGuild As Integer

        Dim GuildTemp As Integer
    
100     With UserList(UserIndex)

102         If .GuildIndex = 0 Then
104             Call WriteConsoleMsg(UserIndex, "¡No perteneces a ningún clan!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
        
106         If .GuildRange <> rFound Then
108             Call WriteConsoleMsg(UserIndex, "¡No eres líder de ningún clan!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
        
110         If Not TieneObjetos(ESCRITURAS_CLAN, 1, UserIndex) Then Exit Sub

        End With
        
        If Not PersonajeExiste(UserName) Then
            Call WriteConsoleMsg(UserIndex, "¡El personaje no existe!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If

112     tUser = NameIndex(UserName)

114     If tUser > 0 Then

116         With UserList(tUser)
            
118             If .GuildIndex <> UserList(UserIndex).GuildIndex Then
120                 Call WriteConsoleMsg(UserIndex, "El personaje debe pertenecer a tu clan.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If

122             SlotGuild = .GuildSlot
124             Call Guilds_Update_Member_Range(tUser, eGuildRange.rFound)

            End With
    
        Else
126         GuildTemp = val(GetVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GUILDINDEX"))

              If GuildTemp > 0 Then
                SlotGuild = Guilds_SlotUser(GuildTemp, UCase$(UserName))
              End If
              
128         If GuildTemp <> UserList(UserIndex).GuildIndex Then
130             Call WriteConsoleMsg(UserIndex, "El personaje debe pertenecer a tu clan.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
132         Call WriteVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GUILDRANGE", CStr(eGuildRange.rFound))

        End If
    
138     Call Guilds_Modify_Member(UserList(UserIndex).GuildIndex, SlotGuild, UserName, eGuildRange.rFound)
    
        ' Chau Old Leader
140     Call Guilds_Update_Member_Range(UserIndex, eGuildRange.rNone)
142     Call Guilds_Modify_Member(UserList(UserIndex).GuildIndex, UserList(UserIndex).GuildSlot, UserList(UserIndex).Name, eGuildRange.rNone)
    
        ' Chau Objs
144     Call QuitarObjetos(ESCRITURAS_CLAN, 1, UserIndex)
146     Call Guilds_Save(UserList(UserIndex).GuildIndex)
    
148     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageConsoleMsg("El líder pasó a ser el personaje " & UserName & " ¡Felicitaciones! Aunque no lo queramos...", FontTypeNames.FONTTYPE_GUILD))
        '<EhFooter>
        Exit Sub

ChangeLeader_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mGuilds.ChangeLeader " & "at line " & Erl

        

        '</EhFooter>
End Sub

