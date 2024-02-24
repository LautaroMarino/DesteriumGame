Attribute VB_Name = "mLog"
' Creado por Lautaro Marino

Option Explicit

Public Enum eLog

    eGeneral = 0
    eUser = 1
    eGm = 2
    eSecurity = 3
    eAccount = 4
End Enum

Public Enum eLogDescUser

    eNone = 0
    
    eDropObj = 1
    eGetObj = 2
    
    eSaleObj = 3
    eBuyObj = 4
    
    eDropGld = 5
    
    eCommerce_Obj = 6
    eCommerce_Gld = 7
    
    eLvl = 8
    
    eBov_Obj = 9
    
    eChat = 10
    eReclameObj = 11
    
    eDropEldhir = 12
    eKill = 13
    eOther = 100
    
End Enum

Public Enum eLogSecurity

    eGeneral = 0
    eAntiCheat = 1
    eAntiFrags = 2
    eAntiHack = 3
    eAutoBan = 4
    eMercader = 5
    eSubastas = 6
    eNewChar = 7
    eShop = 8
    eComerciantes = 9
    eLottery = 10
End Enum

Private Function Logs_FilePath(ByVal Log As eLog) As String
    
    Select Case Log

        Case eLog.eGeneral
            Logs_FilePath = "GENERAL\"

        Case eLog.eUser
            Logs_FilePath = "USER\"

        Case eLog.eGm
            Logs_FilePath = "GM\"

        Case eLog.eSecurity
            Logs_FilePath = "SECURITY\"
            
        Case eLog.eAccount
            Logs_FilePath = "ACCOUNT\"
            
            
    End Select
    
End Function

Public Function Logs_DescUser(ByVal LogDesc As eLogDescUser) As String
    
    Select Case LogDesc

        Case eLogDescUser.eDropObj
            Logs_DescUser = " [DROP OBJ] "

        Case eLogDescUser.eGetObj
            Logs_DescUser = " [GET OBJ] "

        Case eLogDescUser.eSaleObj
            Logs_DescUser = " [SALE OBJ] "

        Case eLogDescUser.eBuyObj
            Logs_DescUser = " [BUY OBJ] "

        Case eLogDescUser.eDropGld
            Logs_DescUser = " [DROP GLD] "
            
        Case eLogDescUser.eDropEldhir
            Logs_DescUser = " [DROP ELDHIR] "
            
        Case eLogDescUser.eKill
            Logs_DescUser = " [FRAG] "
            
        Case eLogDescUser.eCommerce_Obj
            Logs_DescUser = " [COMMERCE OBJ] "

        Case eLogDescUser.eLvl
            Logs_DescUser = " [LEVEL] "

        Case eLogDescUser.eBov_Obj
            Logs_DescUser = " [BOV] "

        Case eLogDescUser.eChat
            Logs_DescUser = " [CHAT] "
            
        Case eLogDescUser.eReclameObj
            Logs_DescUser = " [RETOS OBJ] "
            
        Case Else
            Logs_DescUser = " "
    End Select
    
End Function

Private Function Logs_Security_Path(ByVal LogDesc As eLogSecurity) As String
        Select Case LogDesc

        Case eLogSecurity.eGeneral
            Logs_Security_Path = "GENERAL.log"
            
        Case eLogSecurity.eAntiCheat
            Logs_Security_Path = "ANTICHEAT.log"
            
        Case eLogSecurity.eAntiFrags
            Logs_Security_Path = "ANTIFRAGS.log"
            
        Case eLogSecurity.eAntiHack
            Logs_Security_Path = "ANTIHACK.log"
            
        Case eLogSecurity.eAutoBan
            Logs_Security_Path = "AUTOBAN.log"
            
        Case eLogSecurity.eMercader
            Logs_Security_Path = "MERCADER.log"
            
        Case eLogSecurity.eSubastas
            Logs_Security_Path = "SUBASTAS.log"
            
        Case eLogSecurity.eNewChar
            Logs_Security_Path = "NEWCHARS.log"
            
        Case eLogSecurity.eShop
            Logs_Security_Path = "SHOP.log"
        
        Case eLogSecurity.eComerciantes
            Logs_Security_Path = "COMERCIANTES.log"
            
        Case eLogSecurity.eLottery
            Logs_Security_Path = "LOTTERY.log"
    End Select
    
End Function


Public Sub Logs_User(ByVal UserName As String, _
                     ByVal Log As eLog, _
                     ByVal Desc As eLogDescUser, _
                     ByVal Text As String)
                     
    On Error GoTo ErrHandler
    
    'If Log <> eGm Then
        'Call Logs_Desarrollo(UserName, eLog.eGeneral, Desc, UserName & " " & Text)
    'End If
    
    Dim nfile    As Integer

    Dim FilePath As String
    
    nfile = FreeFile

    FilePath = LogPath & Logs_FilePath(Log) & UCase$(UserName) & ".chr"
    
    
    Open FilePath For Append Shared As #nfile
    Print #nfile, Date & " " & Time & Logs_DescUser(Desc) & " " & Text
    Close #nfile
    
    Logs_Security eLog.eGeneral, eLogSecurity.eGeneral, UCase$(UserName) & " " & Text
    
    Exit Sub

ErrHandler:

End Sub

Public Sub Logs_Desarrollo(ByVal UserName As String, _
                           ByVal Log As eLog, _
                           ByVal Desc As eLogDescUser, _
                           ByVal Text As String)
                     
    On Error GoTo ErrHandler

    Dim nfile    As Integer

    Dim FilePath As String
    
    nfile = FreeFile

    FilePath = LogPath & Logs_FilePath(Log) & "Desarrollo.dat"
    
    Open FilePath For Append Shared As #nfile
    Print #nfile, Date & " " & Time & Logs_DescUser(Desc) & Text
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub



Public Sub Logs_Security(ByVal Log As eLog, _
                         ByVal LogDesc As eLogSecurity, _
                         ByVal Text As String)
                     
    On Error GoTo ErrHandler

    Dim nfile    As Integer

    Dim FilePath As String

    Dim SubPath  As String

    nfile = FreeFile

    'Debug.Print FilePath
    FilePath = LogPath & Logs_FilePath(Log) & Logs_Security_Path(LogDesc)
    
    Open FilePath For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Text
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub


Public Sub Logs_Account_SettingData(ByVal UserIndex As Integer, _
                                    ByVal Tittle As String, _
                                    ByVal Account As String)
                     
    On Error GoTo ErrHandler
    
    Dim nfile    As Integer

    Dim FilePath As String
    
    nfile = FreeFile

    FilePath = LogPath & Logs_FilePath(eLog.eAccount) & UCase$(Account) & ".acc"
    
    Open FilePath For Append Shared As #nfile
    Print #nfile, "[################" & Tittle & " - " & Date & " " & Time & "################]"
    Print #nfile, "IP PUBLICA: " & UserList(UserIndex).Account.Sec.IP_Public
    Print #nfile, "IP ADDRESS: " & UserList(UserIndex).Account.Sec.IP_Address
    Print #nfile, "IP LOCAL: " & UserList(UserIndex).Account.Sec.IP_Local
    Print #nfile, "SERIAL MAC: " & UserList(UserIndex).Account.Sec.SERIAL_MAC
    Print #nfile, "SERIAL DISK: " & UserList(UserIndex).Account.Sec.SERIAL_DISK
    Print #nfile, "SERIAL BIOS: " & UserList(UserIndex).Account.Sec.SERIAL_BIOS
    Print #nfile, "SERIAL MOTHERBOARD: " & UserList(UserIndex).Account.Sec.SERIAL_MOTHERBOARD
    Print #nfile, "SERIAL PROCESSOR: " & UserList(UserIndex).Account.Sec.SERIAL_PROCESSOR
    Print #nfile, "SYSTEM DATA: " & UserList(UserIndex).Account.Sec.SYSTEM_DATA
    Print #nfile, "[################################]"
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub





Public Sub LogCriticEvent(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & Index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEventos(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "EventosDS.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
          
    Exit Sub

ErrHandler:

End Sub

Public Sub LogPerdones(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "Perdones.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogError(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub
Public Sub Log_Reward(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "Reward.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub
Public Sub Log_ChangePjs(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "ChangeInPjs.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogRetos(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "Retos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
          
    Exit Sub

ErrHandler:

End Sub

Public Sub LogStatic(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogTarea(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile(1) ' obtenemos un canal
    Open LogPath & "haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogClanes(ByVal Str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Str
    Close #nfile

End Sub

Public Sub LogIP(ByVal Str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Str
    Close #nfile

End Sub

Public Sub LogAsesinato(Texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open LogPath & "asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub logVentaCasa(ByVal Texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    
    Open LogPath & "propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & Texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogHackAttemp(Texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & Texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogCheating(Texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogCriticalHackAttemp(Texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open LogPath & "CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & Texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

