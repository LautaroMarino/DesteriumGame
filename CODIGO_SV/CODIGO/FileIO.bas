Attribute VB_Name = "ES"
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

Public Sub CargarSpawnList()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CargarSpawnList_Err
        '</EhHeader>

        Dim N As Integer, LoopC As Integer

100     N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
102     ReDim SpawnList(N) As tCriaturasEntrenador

104     For LoopC = 1 To N
106         SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
108         SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
110     Next LoopC
    
        '<EhFooter>
        Exit Sub

CargarSpawnList_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.CargarSpawnList " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        '<EhHeader>
        On Error GoTo EsAdmin_Err
        '</EhHeader>
100     EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)
        '<EhFooter>
        Exit Function

EsAdmin_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.EsAdmin " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function EsDios(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        '<EhHeader>
        On Error GoTo EsDios_Err
        '</EhHeader>
100     EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)
        '<EhFooter>
        Exit Function

EsDios_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.EsDios " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function EsSemiDios(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        '<EhHeader>
        On Error GoTo EsSemiDios_Err
        '</EhHeader>
100     EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)
        '<EhFooter>
        Exit Function

EsSemiDios_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.EsSemiDios " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        '<EhHeader>
        On Error GoTo EsGmEspecial_Err
        '</EhHeader>
100     EsGmEspecial = (val(Administradores.GetValue("Especial", Name)) = 1)
        '<EhFooter>
        Exit Function

EsGmEspecial_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.EsGmEspecial " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/03/2011
        'Returns true if char is administrative user.
        '***************************************************
        '<EhHeader>
        On Error GoTo EsGmChar_Err
        '</EhHeader>
    
        Dim EsGm As Boolean
    
        ' Admin?
100     EsGm = EsAdmin(Name)

        ' Dios?
102     If Not EsGm Then EsGm = EsDios(Name)

        ' Semidios?
104     If Not EsGm Then EsGm = EsSemiDios(Name)

106     EsGmChar = EsGm

        '<EhFooter>
        Exit Function

EsGmChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.EsGmChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub loadAdministrativeUsers()
        'Admines     => Admin
        'Dioses      => Dios
        'SemiDioses  => SemiDios
        'Especiales  => Especial
        '<EhHeader>
        On Error GoTo loadAdministrativeUsers_Err
        '</EhHeader>

        'Si esta mierda tuviese array asociativos el código sería tan lindo.
        Dim buf  As Integer

        Dim i    As Long

        Dim Name As String
       
        ' Public container
100     Set Administradores = New clsIniManager
    
        ' Server ini info file
        Dim ServerIni As clsIniManager

102     Set ServerIni = New clsIniManager
    
104     Call ServerIni.Initialize(IniPath & "Server.ini")
       
        ' Admines
106     buf = val(ServerIni.GetValue("INIT", "Admines"))
    
108     For i = 1 To buf
110         Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
112         If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
            ' Add key
114         Call Administradores.ChangeValue("Admin", Name, "1")

116     Next i
    
        ' Dioses
118     buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
120     For i = 1 To buf
122         Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
124         If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
            ' Add key
126         Call Administradores.ChangeValue("Dios", Name, "1")
        
128     Next i
    
        ' Especiales
130     buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
132     For i = 1 To buf
134         Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
136         If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
            ' Add key
138         Call Administradores.ChangeValue("Especial", Name, "1")
        
140     Next i
    
        ' SemiDioses
142     buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
144     For i = 1 To buf
146         Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
148         If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
            ' Add key
150         Call Administradores.ChangeValue("SemiDios", Name, "1")
        
152     Next i
    
        ' Rangos de GM
154     buf = val(ServerIni.GetValue("RANGOS", "Ultimo"))
    
156     ReDim RangeGm(0 To buf) As tRangeGM

        Dim Temp As String

158     For i = 1 To buf
160         Temp = ServerIni.GetValue("RANGOS", i)
        
162         RangeGm(i).Name = ReadField(1, Temp, Asc("-"))
164         RangeGm(i).Tag = ReadField(2, Temp, Asc("-"))
166     Next i
    
168     Set ServerIni = Nothing
    
        '<EhFooter>
        Exit Sub

loadAdministrativeUsers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.loadAdministrativeUsers " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function GetCharRange(ByVal UserName As String) As String
        '<EhHeader>
        On Error GoTo GetCharRange_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = LBound(RangeGm) To UBound(RangeGm)

102         If RangeGm(A).Name = UserName Then
104             GetCharRange = RangeGm(A).Tag
            End If

106     Next A

        '<EhFooter>
        Exit Function

GetCharRange_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.GetCharRange " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
        '****************************************************
        'Author: ZaMa
        'Last Modification: 18/11/2010
        'Reads the user's charfile and retrieves its privs.
        '***************************************************
        '<EhHeader>
        On Error GoTo GetCharPrivs_Err
        '</EhHeader>

        Dim Privs As PlayerType

100     If EsAdmin(UserName) Then
102         Privs = PlayerType.Admin
        
104     ElseIf EsDios(UserName) Then
106         Privs = PlayerType.Dios

108     ElseIf EsSemiDios(UserName) Then
110         Privs = PlayerType.SemiDios
    
        Else
112         Privs = PlayerType.User
        End If

114     GetCharPrivs = Privs

        '<EhFooter>
        Exit Function

GetCharPrivs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.GetCharPrivs " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function TxtDimension(ByVal Name As String) As Long
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo TxtDimension_Err
        '</EhHeader>

        Dim N As Integer, cad As String, Tam As Long

100     N = FreeFile(1)
102     Open Name For Input As #N
104     Tam = 0

106     Do While Not EOF(N)
108         Tam = Tam + 1
110         Line Input #N, cad
        Loop

112     Close N
114     TxtDimension = Tam
        '<EhFooter>
        Exit Function

TxtDimension_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.TxtDimension " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub CargarForbidenWords()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CargarForbidenWords_Err
        '</EhHeader>

100     ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))

        Dim N As Integer, i As Integer

102     N = FreeFile(1)
104     Open DatPath & "NombresInvalidos.txt" For Input As #N
    
106     For i = 1 To UBound(ForbidenNames)
108         Line Input #N, ForbidenNames(i)
110         ForbidenNames(i) = UCase$(ForbidenNames(i))
112     Next i
    
114     Close N

116     ReDim ForbidenText(1 To TxtDimension(DatPath & "PalabrasInvalidas.txt"))

118     N = FreeFile(1)
120     Open DatPath & "PalabrasInvalidas.txt" For Input As #N
    
122     For i = 1 To UBound(ForbidenText)
124         Line Input #N, ForbidenText(i)
126     Next i
    
128     Close N
        '<EhFooter>
        Exit Sub

CargarForbidenWords_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.CargarForbidenWords " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub



Public Sub CargarHechizos()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    
    Dim Hechizo As Integer

    Dim Leer    As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(Spell_FilePath)
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            .TileRange = val(Leer.GetValue("Hechizo" & Hechizo, "TileRange"))
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
            '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            .SanacionGlobal = val(Leer.GetValue("Hechizo" & Hechizo, "SanacionGlobal"))
            .SanacionGlobalNpcs = val(Leer.GetValue("Hechizo" & Hechizo, "SanacionGlobalNpcs"))
            
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
            '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
            '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            .HpRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "HpRequerido"))
            .AutoLanzar = val(Leer.GetValue("Hechizo" & Hechizo, "AutoLanzar"))
            .AreaX = val(Leer.GetValue("Hechizo" & Hechizo, "AreaX"))
            .AreaY = val(Leer.GetValue("Hechizo" & Hechizo, "AreaY"))
            
            'Barrin 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
            .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
            
            .LvlMin = val(Leer.GetValue("Hechizo" & Hechizo, "LvlMin"))
        End With

    Next Hechizo
    
    Set Leer = Nothing
    
    
    Call DataServer_Generate_Spells
    Exit Sub

ErrHandler:
    MsgBox "Error cargando " & Spell_FilePath & " " & Err.number & ": " & Err.description
 
End Sub

Sub LoadMotd()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LoadMotd_Err
        '</EhHeader>

        Dim i As Integer
    
100     MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
    
102     ReDim MOTD(1 To MaxLines)

104     For i = 1 To MaxLines
106         MOTD(i).Texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
108         MOTD(i).Formato = vbNullString
110     Next i

        '<EhFooter>
        Exit Sub

LoadMotd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadMotd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoBackUp()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo DoBackUp_Err
        '</EhHeader>

100     haciendoBK = True
    
102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
104     Call LimpiarMundo
106     Call WorldSave

108     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
        'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)
    
110     haciendoBK = False
    

        Dim nfile As Integer

112     nfile = FreeFile ' obtenemos un canal
114     Open LogPath & "BackUps.log" For Append Shared As #nfile
116     Print #nfile, Date & " " & Time
118     Close #nfile
        '<EhFooter>
        Exit Sub

DoBackUp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.DoBackUp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2011
        '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
        '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
        '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados)
        '***************************************************
        '<EhHeader>
        On Error GoTo GrabarMapa_Err
        '</EhHeader>

        Dim FreeFileMap As Long

        Dim FreeFileInf As Long

        Dim Y           As Long

        Dim X           As Long

        Dim ByFlags     As Byte

        Dim LoopC       As Long

        Dim MapWriter   As clsByteBuffer

        Dim InfWriter   As clsByteBuffer

        Dim IniManager  As clsIniManager

        Dim NpcInvalido As Boolean
    
100     Set MapWriter = New clsByteBuffer
102     Set InfWriter = New clsByteBuffer
104     Set IniManager = New clsIniManager
    
106     If FileExist(MAPFILE & ".map", vbNormal) Then
108         Kill MAPFILE & ".map"
        End If
    
110     If FileExist(MAPFILE & ".inf", vbNormal) Then
112         Kill MAPFILE & ".inf"
        End If
    
        'Open .map file
114     FreeFileMap = FreeFile
116     Open MAPFILE & ".map" For Binary As FreeFileMap
    
118     Call MapWriter.initializeWriter(FreeFileMap)
    
        'Open .inf file
120     FreeFileInf = FreeFile
122     Open MAPFILE & ".inf" For Binary As FreeFileInf
    
124     Call InfWriter.initializeWriter(FreeFileInf)
    
        'map Header
126     Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
128     Call MapWriter.putString(MiCabecera.Desc, False)
130     Call MapWriter.putLong(MiCabecera.CRC)
132     Call MapWriter.putLong(MiCabecera.MagicWord)
    
134     Call MapWriter.putDouble(0)
    
        'inf Header
136     Call InfWriter.putDouble(0)
138     Call InfWriter.putInteger(0)
    
        'Write .map file
140     For Y = YMinMapSize To YMaxMapSize
142         For X = XMinMapSize To XMaxMapSize

144             With MapData(Map, X, Y)
146                 ByFlags = 0
                
148                 If .Blocked Then ByFlags = ByFlags Or 1
150                 If .Graphic(2) Then ByFlags = ByFlags Or 2
152                 If .Graphic(3) Then ByFlags = ByFlags Or 4
154                 If .Graphic(4) Then ByFlags = ByFlags Or 8
156                 If .trigger Then ByFlags = ByFlags Or 16
                
158                 Call MapWriter.putByte(ByFlags)
                
160                 Call MapWriter.putLong(CLng(.Graphic(1)))
                
162                 For LoopC = 2 To 4

164                     If .Graphic(LoopC) Then Call MapWriter.putLong(CLng(.Graphic(LoopC)))
166                 Next LoopC
                
168                 If .trigger Then Call MapWriter.putInteger(CInt(.trigger))
                
                    '.inf file
170                 ByFlags = 0
                
172                 If .ObjInfo.ObjIndex > 0 Then
174                     If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
176                         .ObjInfo.ObjIndex = 0
178                         .ObjInfo.Amount = 0
                        End If
                    End If
    
180                 If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                    ' No hacer backup de los NPCs inválidos (Pretorianos, Mascotas, Invocados, npcs invocados)
182                 If .NpcIndex Then
184                     NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0) Or (Npclist(.NpcIndex).flags.Invocation = 1)
                    
186                     If Not NpcInvalido Then ByFlags = ByFlags Or 2
                    End If
                
188                 If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
190                 Call InfWriter.putByte(ByFlags)
                
192                 If .TileExit.Map Then
194                     Call InfWriter.putInteger(.TileExit.Map)
196                     Call InfWriter.putInteger(.TileExit.X)
198                     Call InfWriter.putInteger(.TileExit.Y)
                    End If
                
200                 If .NpcIndex And Not NpcInvalido Then Call InfWriter.putInteger(Npclist(.NpcIndex).numero)
                
202                 If .ObjInfo.ObjIndex Then
204                     Call InfWriter.putInteger(.ObjInfo.ObjIndex)
206                     Call InfWriter.putInteger(.ObjInfo.Amount)
                    End If
                
208                 NpcInvalido = False
                End With

210         Next X
212     Next Y
    
214     Call MapWriter.saveBuffer
216     Call InfWriter.saveBuffer
    
        'Close .map file
218     Close FreeFileMap

        'Close .inf file
220     Close FreeFileInf
    
222     Set MapWriter = Nothing
224     Set InfWriter = Nothing

226     With MapInfo(Map)
            'write .dat file
228         Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
230         Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .Music)
232         Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
234         Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
236         Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
238         Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
240         Call IniManager.ChangeValue("Mapa" & Map, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y)
242         Call IniManager.ChangeValue("Mapa" & Map, "OnLoginGoTo", .OnLoginGoTo.Map & "-" & .OnLoginGoTo.X & "-" & .OnLoginGoTo.Y)
            Call IniManager.ChangeValue("Mapa" & Map, "GoToOns", .GoToOns.Map & "-" & .GoToOns.X & "-" & .GoToOns.Y)
244         Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainByteToString(.Terreno))
246         Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
248         Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.Restringir))
250         Call IniManager.ChangeValue("Mapa" & Map, "BackUp", Str(.BackUp))
    
252         If .Pk Then
254             Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
            Else
256             Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")
            End If
        
258         Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
260         Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
262         Call IniManager.ChangeValue("Mapa" & Map, "MimetismoSinEfecto", .MimetismoSinEfecto)
              Call IniManager.ChangeValue("Mapa" & Map, "Faction", .Faction)
264         Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
266         Call IniManager.ChangeValue("Mapa" & Map, "LvlMin", .LvlMin)
268         Call IniManager.ChangeValue("Mapa" & Map, "LvlMax", .LvlMax)
270         Call IniManager.ChangeValue("Mapa" & Map, "Premium", .Premium)
272         Call IniManager.ChangeValue("Mapa" & Map, "Limpieza", .Limpieza)
274         Call IniManager.ChangeValue("Mapa" & Map, "CaenItems", .CaenItems)
276         Call IniManager.ChangeValue("Mapa" & Map, "Bronce", .Bronce)
278         Call IniManager.ChangeValue("Mapa" & Map, "Plata", .Plata)
280         Call IniManager.ChangeValue("Mapa" & Map, "Guild", .Guild)
            Call IniManager.ChangeValue("Mapa" & Map, "NOMANA", CStr(.NoMana))
            
            Call IniManager.ChangeValue("Mapa" & Map, "Poder", CStr(.Poder))
                
282         Call IniManager.DumpFile(MAPFILE & ".dat")
        End With
    
284     Set IniManager = Nothing
        '<EhFooter>
        Exit Sub

GrabarMapa_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.GrabarMapa " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LoadBalance()
        '***************************************************
        'Author: Unknown
        'Last Modification: 15/04/2010
        '15/04/2010: ZaMa - Agrego recompensas faccionarias.
        '***************************************************
        '<EhHeader>
        On Error GoTo LoadBalance_Err
        '</EhHeader>

        Dim i As Long, A As Long
    
        Dim FilePath As String
        Dim Prefix As String
        Dim Temp As String
        Dim Arrai() As String
    
100     FilePath = DatPath & "Balance.dat"
102     Prefix = "-level"
    
104     With Balance
        
            ' Configuración inicial [INIT]
           ' .PASSIVE_MAX = val(GetVar(FilePath, "INIT", "PASSIVE_MAX"))         ' Máximo valor de pasiva según la clase
        
            'Modificadores de Clase
106         For i = 1 To NUMCLASES
            
              '  Temp = GetVar(FilePath, "INV_CLASS", ListaClases(i))
              '  Arrai = Split(Temp, "-")
            
              '  ReDim .ListObjs(i).Obj(LBound(Arrai) To UBound(Arrai)) As Integer
            '    For A = LBound(Arrai) To UBound(Arrai)
                 '   .ListObjs(i).Obj(A) = val(Arrai(A))
              '  Next A
            
             '   .RazeClass(i) = val(GetVar(FilePath, "RAZE_CLASS", ListaClases(i)))
                '.GeneroClass(i) = val(GetVar(FilePath, "GENERO_CLASS", ListaClases(i)))
            
                '.Health_Initial(i) = val(GetVar(FilePath, "HEALTH", ListaClases(i)))
              '  .Health_Level(i) = val(GetVar(FilePath, "HEALTH", ListaClases(i) & Prefix))
            
              '  .Mana_Initial(i) = val(GetVar(FilePath, "MANA", ListaClases(i)))
             '   .Mana_Level(i) = val(GetVar(FilePath, "MANA", ListaClases(i) & Prefix))
            
              '  .Damage_Initial(i) = val(GetVar(FilePath, "DAMAGE", ListaClases(i)))
              '  .Damage_Level(i) = val(GetVar(FilePath, "DAMAGE", ListaClases(i) & Prefix))
            
              '  .DamageMag_Initial(i) = val(GetVar(FilePath, "DAMAGE_MAGIC", ListaClases(i)))
             '   .DamageMag_Level(i) = val(GetVar(FilePath, "DAMAGE_MAGIC", ListaClases(i) & Prefix))
            
               ' .Armour_Initial(i) = val(GetVar(FilePath, "ARMOUR", ListaClases(i)))
               ' .Armour_Level(i) = val(GetVar(FilePath, "ARMOUR", ListaClases(i) & Prefix))
            
               ' .ArmourMag_Initial(i) = val(GetVar(FilePath, "ARMOUR_MAGIC", ListaClases(i)))
               ' .ArmourMag_Level(i) = val(GetVar(FilePath, "ARMOUR_MAGIC", ListaClases(i) & Prefix))
            
              '  .Attack_Initial(i) = val(GetVar(FilePath, "ATTACK", ListaClases(i)))
               ' .Attack_Level(i) = val(GetVar(FilePath, "ATTACK", ListaClases(i) & Prefix))
            
              '  .RegHP_Initial(i) = val(GetVar(FilePath, "REGENERATION_HP", ListaClases(i)))
              '  .RegHP_Level(i) = val(GetVar(FilePath, "REGENERATION_HP", ListaClases(i) & Prefix))
            
              '  .RegMANA_Initial(i) = val(GetVar(FilePath, "REGENERATION_MANA", ListaClases(i)))
              '  .RegMANA_Level(i) = val(GetVar(FilePath, "REGENERATION_MANA", ListaClases(i) & Prefix))
            
                '.Movement_Initial(i) = val(GetVar(FilePath, "VELOCITY_CHAR", ListaClases(i)))
            
              '  .Cooldown_Initial(i) = val(GetVar(FilePath, "COOLDOWN", ListaClases(i)))
               ' .Cooldown_Level(i) = val(GetVar(FilePath, "COOLDOWN", ListaClases(i) & Prefix))
            
108             With .ModClase(i)
110                 .Evasion = val(GetVar(FilePath, "MODEVASION", ListaClases(i)))
112                 .AtaqueArmas = val(GetVar(FilePath, "MODATAQUEARMAS", ListaClases(i)))
114                 .AtaqueProyectiles = val(GetVar(FilePath, "MODATAQUEPROYECTILES", ListaClases(i)))
116                 .AtaqueWrestling = val(GetVar(FilePath, "MODATAQUEWRESTLING", ListaClases(i)))
118                 .DañoArmas = val(GetVar(FilePath, "MODDAÑOARMAS", ListaClases(i)))
120                 .DañoProyectiles = val(GetVar(FilePath, "MODDAÑOPROYECTILES", ListaClases(i)))
122                 .DañoWrestling = val(GetVar(FilePath, "MODDAÑOWRESTLING", ListaClases(i)))
124                 .Escudo = val(GetVar(FilePath, "MODESCUDO", ListaClases(i)))

                End With

126         Next i
    
            'Modificadores de Raza
128         For i = 1 To NUMRAZAS

130             With .ModRaza(i)
132                 .Fuerza = val(GetVar(FilePath, "MODRAZA", ListaRazas(i) + "Fuerza"))
134                 .Agilidad = val(GetVar(FilePath, "MODRAZA", ListaRazas(i) + "Agilidad"))
136                 .Inteligencia = val(GetVar(FilePath, "MODRAZA", ListaRazas(i) + "Inteligencia"))
138                 .Carisma = val(GetVar(FilePath, "MODRAZA", ListaRazas(i) + "Carisma"))
140                 .Constitucion = val(GetVar(FilePath, "MODRAZA", ListaRazas(i) + "Constitucion"))

                End With

142         Next i
    
            'Modificadores de Vida
144         For i = 1 To NUMCLASES
146             .ModVida(i) = val(GetVar(FilePath, "MODVIDA", ListaClases(i)))
148         Next i
    
            'Distribución de Vida
150         For i = 1 To 5
152             .DistribucionEnteraVida(i) = val(GetVar(FilePath, "DISTRIBUCION", "E" + CStr(i)))
154         Next i

156         For i = 1 To 4
158             .DistribucionSemienteraVida(i) = val(GetVar(FilePath, "DISTRIBUCION", "S" + CStr(i)))
160         Next i
    
            'Extra
162         .PorcentajeRecuperoMana = val(GetVar(FilePath, "EXTRA", "PorcentajeRecuperoMana"))
        
            ' Recompensas faccionarias
            'For i = 1 To NUM_RANGOS_FACCION
            'RecompensaFacciones(i - 1) = val(GetVar(Filepath, "RECOMPENSAFACCION", "Rango" & i))
            'Next i
        
        End With
    
        '<EhFooter>
        Exit Sub

LoadBalance_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadBalance " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub LoadOBJData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer

    Dim A      As Long

    Dim Temp   As String
    
    Dim Leer   As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(Objs_FilePath)
    
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    'Llena la lista
    For Object = 1 To NumObjDatas

        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            .NoNada = val(Leer.GetValue("OBJ" & Object, "NONADA"))
            .NoDrop = val(Leer.GetValue("OBJ" & Object, "NoDrop"))

            'Pablo (ToxicWaste) Log de Objetos.
            .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            
            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex

            End If
            
            .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
            .Donable = val(Leer.GetValue("OBJ" & Object, "Donable"))
            .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            .Bronce = val(Leer.GetValue("OBJ" & Object, "Bronce"))
            .Plata = val(Leer.GetValue("OBJ" & Object, "Plata"))
            .Oro = val(Leer.GetValue("OBJ" & Object, "Oro"))
            .Premium = val(Leer.GetValue("OBJ" & Object, "Premium"))
            
            .MagiaSkill = val(Leer.GetValue("OBJ" & Object, "MagiaSkill"))
            .RMSkill = val(Leer.GetValue("OBJ" & Object, "RMSkill"))
            .ArmaSkill = val(Leer.GetValue("OBJ" & Object, "WeaponSkill"))
            .EscudoSkill = val(Leer.GetValue("OBJ" & Object, "EscudoSkill"))
            .ArmaduraSkill = val(Leer.GetValue("OBJ" & Object, "ArmaduraSkill"))
            .ArcoSkill = val(Leer.GetValue("OBJ" & Object, "ArcoSkill"))
            .DagaSkill = val(Leer.GetValue("OBJ" & Object, "DagaSkill"))
            .QuitaEnergia = val(Leer.GetValue("OBJ" & Object, "Energia"))
            .EdicionLimitada = val(Leer.GetValue("OBJ" & Object, "EdicionLimitada"))
            .Navidad = val(Leer.GetValue("OBJ" & Object, "Navidad"))
            .Ilimitado = val(Leer.GetValue("OBJ" & Object, "Ilimitado"))
            .Sound = val(Leer.GetValue("OBJ" & Object, "Sound"))
            .DosManos = val(Leer.GetValue("OBJ" & Object, "DosManos"))
            
            
            
            Select Case .OBJType
                
                Case eOBJType.otMagic
                    .StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                    
                Case eOBJType.otUseOnce
                    .ProbPesca = val(Leer.GetValue("OBJ" & Object, "ProbPesca"))
                    
                Case eOBJType.otarmadura
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .AuraIndex(1) = val(Leer.GetValue("OBJ" & Object, "AuraIndex"))

                    .Oculto = val(Leer.GetValue("OBJ" & Object, "Oculto"))
                    
                Case eOBJType.otescudo
                    .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .NoShield = val(Leer.GetValue("OBJ" & Object, "NoShield"))
                    .AuraIndex(4) = val(Leer.GetValue("OBJ" & Object, "AuraIndex"))
                    
                Case eOBJType.otcasco
                    .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .Oculto = val(Leer.GetValue("OBJ" & Object, "Oculto"))      ' capuchas
                    .AuraIndex(3) = val(Leer.GetValue("OBJ" & Object, "AuraIndex"))
                    
                Case eOBJType.otPendienteParty
                    .Porc = val(Leer.GetValue("OBJ" & Object, "Porc"))
                    
                Case eOBJType.otWeapon
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                    .StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                    .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    .NpcBonusDamage = val(Leer.GetValue("OBJ" & Object, "NpcBonusDamage"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                    
                    .ProbPesca = val(Leer.GetValue("OBJ" & Object, "ProbPesca"))
                    .AuraIndex(2) = val(Leer.GetValue("OBJ" & Object, "AuraIndex"))
                    
                Case eOBJType.otAuras
                    
                    
                Case eOBJType.oteffect
                    .RemoveObj = val(Leer.GetValue("OBJ" & Object, "RemoveObj"))
                    .Time = val(Leer.GetValue("OBJ" & Object, "Time"))
                    .BonusTipe = val(Leer.GetValue("OBJ" & Object, "BonusTipe"))
                    .BonusValue = val(Leer.GetValue("OBJ" & Object, "BonusValue"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    
                    
                    
                Case eOBJType.otTransformVIP
                    .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
                
                Case eOBJType.otTravel
                    .TelepMap = val(Leer.GetValue("OBJ" & Object, "TelepMap"))
                    .TelepX = val(Leer.GetValue("OBJ" & Object, "TelepX"))
                    .TelepY = val(Leer.GetValue("OBJ" & Object, "TelepY"))
                    .RequiredNpc = val(Leer.GetValue("OBJ" & Object, "NpcNumber"))
                
                Case eOBJType.otLibroGuild
                    .GuildExp = val(Leer.GetValue("OBJ" & Object, "GuildExp"))
                    
                Case eOBJType.otReliquias
                    .EffectUser.Hp = val(Leer.GetValue("OBJ" & Object, "Hp"))
                    .EffectUser.Man = val(Leer.GetValue("OBJ" & Object, "Man"))
                    .EffectUser.Damage = val(Leer.GetValue("OBJ" & Object, "Damage"))
                    .EffectUser.DamageMagic = val(Leer.GetValue("OBJ" & Object, "DamageMagic"))
                    .EffectUser.DamageNpc = val(Leer.GetValue("OBJ" & Object, "DamageNpc"))
                    .EffectUser.DamageMagicNpc = val(Leer.GetValue("OBJ" & Object, "DamageMagicNc"))
                    .EffectUser.AfectaParalisis = val(Leer.GetValue("OBJ" & Object, "AfectaParalisis"))
                    .EffectUser.DevuelveVidaPorc = val(Leer.GetValue("OBJ" & Object, "DevuelveVida"))
                    .EffectUser.ExpNpc = val(Leer.GetValue("OBJ" & Object, "ExpNpc"))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))

                Case eOBJType.otItemRandom
                    .MaxFortunas = val(Leer.GetValue("OBJ" & Object, "NROITEMS"))
                     
                    If .MaxFortunas > 0 Then
                        ReDim .Fortuna(1 To .MaxFortunas) As Obj
                        
                        For A = 1 To .MaxFortunas
                            Temp = Leer.GetValue("OBJ" & Object, "OBJ" & A)
                            
                            .Fortuna(A).ObjIndex = val(ReadField(1, Temp, 45))
                            .Fortuna(A).Amount = val(ReadField(2, Temp, 45))
                        Next A

                    End If
                    
                Case eOBJType.otCofreAbierto
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                     
                    ' Cofres Cerrados
                Case eOBJType.otcofre
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .Chest.NroDrop = val(Leer.GetValue("OBJ" & Object, "NroDrop"))
                    
                    If .Chest.NroDrop > 0 Then
                        ReDim .Chest.Drop(1 To .Chest.NroDrop) As Integer
                        
                        Temp = Leer.GetValue("OBJ" & Object, "Drops")
                        
                        For A = 1 To .Chest.NroDrop
                            .Chest.Drop(A) = val(ReadField(A, Temp, 45))
                        Next A

                    End If

                    .Chest.RespawnTime = val(Leer.GetValue("OBJ" & Object, "RespawnTime"))
                    .Chest.ProbClose = val(Leer.GetValue("OBJ" & Object, "ProbClose"))
                    .Chest.ProbBreak = val(Leer.GetValue("OBJ" & Object, "ProbBreak"))
                    .Chest.ClicTime = val(Leer.GetValue("OBJ" & Object, "ClicTime"))
                    
                Case eOBJType.otGemasEffect
                    .BonoExp = val(Leer.GetValue("OBJ" & Object, "BonoExp"))
                    .BonoGld = val(Leer.GetValue("OBJ" & Object, "BonoGld"))
                    .BonoEvasion = val(Leer.GetValue("OBJ" & Object, "BonoEvasion"))
                    .BonoRm = val(Leer.GetValue("OBJ" & Object, "BonoRm"))
                    .BonoArcos = val(Leer.GetValue("OBJ" & Object, "BonoArcos"))
                    .BonoArmas = val(Leer.GetValue("OBJ" & Object, "BonoArmas"))
                    .BonoHechizos = val(Leer.GetValue("OBJ" & Object, "BonoHechizos"))
                    .BonoTime = val(Leer.GetValue("OBJ" & Object, "BonoTime"))
                    
                Case eOBJType.otGemaTelep
                    .TelepMap = val(Leer.GetValue("OBJ" & Object, "TelepMap"))
                    .TelepX = val(Leer.GetValue("OBJ" & Object, "TelepX"))
                    .TelepY = val(Leer.GetValue("OBJ" & Object, "TelepY"))
                    .TelepTime = val(Leer.GetValue("OBJ" & Object, "TelepTime"))
                
                Case otPociones
                    .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otRangeQuest
                    .Range = val(Leer.GetValue("OBJ" & Object, "Range"))
                    
                Case eOBJType.otBarcos
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .ProbPesca = val(Leer.GetValue("OBJ" & Object, "ProbPesca"))
                    
                Case eOBJType.otFlechas
                    .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    .VictimAnim = val(Leer.GetValue("OBJ" & Object, "VictimAnim"))
                    .Incineracion = val(Leer.GetValue("OBJ" & Object, "Incineracion"))
                    
                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    
                Case eOBJType.otTeleport
                    .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                    .FX = val(Leer.GetValue("OBJ" & Object, "Anim"))
                     
                Case eOBJType.otMochilas
                    .MochilaType = val(Leer.GetValue("OBJ" & Object, "MochilaType"))
                    
                Case eOBJType.otForos

                    ' Menues desplegables p/objeto
                Case eOBJType.otYunque
                    .MenuIndex = eMenues.ieYunque
                    
                Case eOBJType.otFragua
                    .MenuIndex = eMenues.ieFragua

                Case eOBJType.otTeleportInvoker
                    .RemoveObj = val(Leer.GetValue("OBJ" & Object, "RemoveObj"))
                    .TimeWarp = val(Leer.GetValue("OBJ" & Object, "TimeWarp"))
                    .TimeDuration = val(Leer.GetValue("OBJ" & Object, "TimeDuration"))
                    
                    Temp = Leer.GetValue("OBJ" & Object, "Position")
                    .Position.Map = val(ReadField(1, Temp, 45))
                    .Position.X = val(ReadField(2, Temp, 45))
                    .Position.Y = val(ReadField(3, Temp, 45))
                    
                    .TeleportObj = val(Leer.GetValue("OBJ" & Object, "TeleportObj"))
                    .FX = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .PuedeInsegura = val(Leer.GetValue("OBJ" & Object, "PuedeInsegura"))

            End Select
            
            ' Menues desplegables p/objeto
            If Object = Leña Or Object = LeñaTejo Or Object = LeñaRoble Then
                .MenuIndex = eMenues.ieLenia
            ElseIf Object = FOGATA Then
                .MenuIndex = eMenues.ieFogata
            ElseIf Object = FOGATA_APAG Then
                .MenuIndex = eMenues.ieRamas

            End If
            
            .RopajeEnano = val(Leer.GetValue("OBJ" & Object, "RopajeEnano"))
            .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .Valor = IIf(.OBJType = eOBJType.otItemNpc, val(Leer.GetValue("OBJ" & Object, "Valor")) * MultGld, val(Leer.GetValue("OBJ" & Object, "Valor")))
            .ValorEldhir = val(Leer.GetValue("OBJ" & Object, "ValorEldhir"))
            .Tier = val(Leer.GetValue("OBJ" & Object, "Tier"))
            .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

            End If
            
            'Puertas y llaves
            .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .Texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
            .Guante = val(Leer.GetValue("OBJ" & Object, "Guante"))
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            Dim i As Integer

            Dim N As Integer

            Dim S As String

            For i = 1 To NUMCLASES
                S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
                N = 1

                Do While LenB(S) > 0 And S <> "0" And UCase$(ListaClases(N)) <> S
                    N = N + 1
                Loop

                .ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
            Next i
            
            .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))

            'Bebidas
            .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            .ArbolItem = val(Leer.GetValue("OBJ" & Object, "ArbolItem"))
            .Points = val(Leer.GetValue("OBJ" & Object, "Points"))
            .LvlMax = val(Leer.GetValue("OBJ" & Object, "LvlMax"))
            .LvlMin = val(Leer.GetValue("OBJ" & Object, "LvlMin"))
            
            .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            .Upgrade.RequiredPremium = val(Leer.GetValue("OBJ" & Object, "REQUIREDPREMIUM"))
            
            'If .Upgrade.RequiredPremium > 0 Then
            'Call mWork.Crafting_BlackSmith_Add(Object, True)
            'End If
            
            .Upgrade.RequiredCant = val(Leer.GetValue("OBJ" & Object, "REQUIREDCANT"))

            If .Upgrade.RequiredCant > 0 Then
                ReDim .Upgrade.Required(1 To .Upgrade.RequiredCant) As Obj
                    
                For A = 1 To .Upgrade.RequiredCant
                    Temp = Leer.GetValue("OBJ" & Object, "R" & A)
                    .Upgrade.Required(A).ObjIndex = val(ReadField(1, Temp, Asc("-")))
                    .Upgrade.Required(A).Amount = val(ReadField(2, Temp, Asc("-")))
                Next A

            End If
            
            .SizeWidth = CByte(val(Leer.GetValue("OBJ" & Object, "SizeWidth")))
            .SizeHeight = CByte(val(Leer.GetValue("OBJ" & Object, "SizeHeight")))
            
            If .SizeWidth = 0 Then .SizeWidth = ModAreas.DEFAULT_ENTITY_WIDTH
            If .SizeHeight = 0 Then .SizeHeight = ModAreas.DEFAULT_ENTITY_HEIGHT
            
            .Skin = CByte(val(Leer.GetValue("OBJ" & Object, "Skin")))
            .GuildLvl = CByte(val(Leer.GetValue("OBJ" & Object, "GuildLvl")))
            
            .Dead = CByte(val(Leer.GetValue("OBJ" & Object, "Dead")))
            .DurationDay = CInt(val(Leer.GetValue("OBJ" & Object, "DurationDays")))
            .AntiFrio = CByte(val(Leer.GetValue("OBJ" & Object, "AntiFrio")))
            
            ' Skills/Atributos de los objetos equipables.
            .SkillNum = CByte(val(Leer.GetValue("OBJ" & Object, "Skills")))
            
            If .SkillNum > 0 Then
                ReDim .Skill(1 To .SkillNum) As ObjData_Skills
            
                For A = 1 To .SkillNum
                    S = Leer.GetValue("OBJ" & Object, "Sk" & A)
                    .Skill(A).Selected = val(ReadField(1, S, 45))
                    .Skill(A).Amount = val(ReadField(2, S, 45))
                Next A
            
            End If
            
            ' Skills/Atributos especiales
            .SkillsEspecialNum = CByte(val(Leer.GetValue("OBJ" & Object, "SkillsEspeciales")))
            
            If .SkillsEspecialNum > 0 Then
                ReDim .SkillsEspecial(1 To .SkillsEspecialNum) As ObjData_Skills
            
                For A = 1 To .SkillsEspecialNum
                    S = Leer.GetValue("OBJ" & Object, "SkEsp" & A)
                    .SkillsEspecial(A).Selected = val(ReadField(1, S, 45))
                    .SkillsEspecial(A).Amount = val(ReadField(2, S, 45))
                Next A

            End If

            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
          '  If .Skin > 0 Then
              '  If .ValorEldhir > 0 Then
                    '.ValorEldhir = Int(.ValorEldhir / 4)
                   ' Call WriteVar(Objs_FilePath, "OBJ" & Object, "ValorEldhir", CStr(.ValorEldhir))
                'End If
           ' End If
        End With

    Next Object
    
    Set Leer = Nothing
    
    
    
    Exit Sub

ErrHandler:
    MsgBox "error cargando objetos " & Err.number & ": " & Err.description

End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)
        '<EhHeader>
        On Error GoTo LoadUserStats_Err
        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 11/19/2009
        '11/19/2009: Pato - Load the EluSkills and ExpSkills
        '*************************************************
        Dim LoopC As Long
        Dim A As Long
        Dim Temp As String
        
100     With UserList(UserIndex)
102         With .Stats

104             For LoopC = 1 To NUMATRIBUTOS
106                 .UserAtributos(LoopC) = CInt(Userfile.GetValue("ATRIBUTOS", "AT" & LoopC))
108                 .UserAtributosBackUP(LoopC) = .UserAtributos(LoopC)
110             Next LoopC
        
                For LoopC = 1 To NUMSKILLSESPECIAL
                      .UserSkillsEspecial(LoopC) = val(Userfile.GetValue("SKILLSESPECIAL", "SKESP" & LoopC))
                Next LoopC
                
112             For LoopC = 1 To NUMSKILLS
                   
114                 .UserSkills(LoopC) = val(Userfile.GetValue("SKILLS", "SK" & LoopC))
116                 .EluSkills(LoopC) = val(Userfile.GetValue("SKILLS", "ELUSK" & LoopC))
118                 .ExpSkills(LoopC) = val(Userfile.GetValue("SKILLS", "EXPSK" & LoopC))
120             Next LoopC
        
122             For LoopC = 1 To MAXUSERHECHIZOS
124                 .UserHechizos(LoopC) = val(Userfile.GetValue("Hechizos", "H" & LoopC))
126             Next LoopC
                
                
                .BonusTipe = val(Userfile.GetValue("STATS", "BONUSTIPE"))
                .BonusValue = CSng(Userfile.GetValue("STATS", "BONUSVALUE"))
                
                'RANKING PERSONAL
128             '.Retos1Jugados = CLng(UserFile.GetValue("RANKING", "RETOS1JUGADOS"))
130             '.Retos1Ganados = CLng(UserFile.GetValue("RANKING", "RETOS1GANADOS"))
132             '.DesafiosJugados = CLng(UserFile.GetValue("RANKING", "RETOS2JUGADOS"))
134             '.DesafiosGanados = CLng(UserFile.GetValue("RANKING", "RETOS2GANADOS"))
136             '.TorneosJugados = CLng(UserFile.GetValue("RANKING", "TORNEOSJUGADOS"))
138             '.TorneosGanados = CLng(UserFile.GetValue("RANKING", "TORNEOSGANADOS"))
        
140             .Eldhir = CLng(Userfile.GetValue("STATS", "ELDHIR"))
142             .BonosHp = CLng(Userfile.GetValue("STATS", "BONOSHP"))
144             .Gld = CLng(Userfile.GetValue("STATS", "GLD"))
        
146             .MaxHp = CInt(Userfile.GetValue("STATS", "MaxHP"))
148             .MinHp = CInt(Userfile.GetValue("STATS", "MinHP"))
        
150             .MinSta = CInt(Userfile.GetValue("STATS", "MinSTA"))
152             .MaxSta = CInt(Userfile.GetValue("STATS", "MaxSTA"))
        
154             .MaxMan = CInt(Userfile.GetValue("STATS", "MaxMAN"))
156             .MinMan = CInt(Userfile.GetValue("STATS", "MinMAN"))
        
158             .MaxHit = CInt(Userfile.GetValue("STATS", "MaxHIT"))
160             .MinHit = CInt(Userfile.GetValue("STATS", "MinHIT"))
        
162             .MaxAGU = CByte(Userfile.GetValue("STATS", "MaxAGU"))
164             .MinAGU = CByte(Userfile.GetValue("STATS", "MinAGU"))
        
166             .MaxHam = CByte(Userfile.GetValue("STATS", "MaxHAM"))
168             .MinHam = CByte(Userfile.GetValue("STATS", "MinHAM"))
        
170             .SkillPts = CInt(Userfile.GetValue("STATS", "SkillPtsLibres"))
        
172             .Exp = CDbl(Userfile.GetValue("STATS", "EXP"))
174             .Elu = CLng(Userfile.GetValue("STATS", "ELU"))
176             .Elv = CByte(Userfile.GetValue("STATS", "ELV"))
                
                
                .BonusLast = CInt(Userfile.GetValue("BONUS", "BONUSLAST"))
                
                If .BonusLast > 0 Then
                ReDim .Bonus(1 To .BonusLast) As UserBonus
                
                For A = 1 To .BonusLast
                    Temp = Userfile.GetValue("BONUS", "BONUS" & A)
                    .Bonus(A).Tipo = val(ReadField(1, Temp, Asc("|")))
                    .Bonus(A).Value = val(ReadField(2, Temp, Asc("|")))
                    .Bonus(A).Amount = val(ReadField(3, Temp, Asc("|")))
                    .Bonus(A).DurationSeconds = val(ReadField(4, Temp, Asc("|")))
                    .Bonus(A).DurationDate = ReadField(5, Temp, Asc("|"))
                Next A
                End If
182             '.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
184             .NPCsMuertos = CInt(Userfile.GetValue("MUERTES", "NpcsMuertes"))
            End With
    
186         With .flags

188             If CByte(Userfile.GetValue("CONSEJO", "PERTENECE")) Then .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
190             If CByte(Userfile.GetValue("CONSEJO", "PERTENECECAOS")) Then .Privilegios = .Privilegios Or PlayerType.ChaosCouncil
            End With
        End With

        '<EhFooter>
        Exit Sub

LoadUserStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadUserStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LoadUserMeditations(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)
        '<EhHeader>
        On Error GoTo LoadUserMeditations_Err
        '</EhHeader>
    
        Dim A As Long
    
100     With UserList(UserIndex)

102         For A = 1 To MAX_MEDITATION
104             .MeditationUser(A) = val(Userfile.GetValue("MEDITATION", A))
106         Next A
    
108         .MeditationSelected = val(Userfile.GetValue("MEDITATION", "SELECTED"))
        End With

        '<EhFooter>
        Exit Sub

LoadUserMeditations_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadUserMeditations " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LoadUserReputacion_Err
        '</EhHeader>

100     With UserList(UserIndex).Reputacion
102         .AsesinoRep = val(Userfile.GetValue("REP", "Asesino"))
104         .BandidoRep = val(Userfile.GetValue("REP", "Bandido"))
106         .BurguesRep = val(Userfile.GetValue("REP", "Burguesia"))
108         .LadronesRep = val(Userfile.GetValue("REP", "Ladrones"))
110         .NobleRep = val(Userfile.GetValue("REP", "Nobles"))
112         .PlebeRep = val(Userfile.GetValue("REP", "Plebe"))
114         .promedio = val(Userfile.GetValue("REP", "Promedio"))
        End With
    
        '<EhFooter>
        Exit Sub

LoadUserReputacion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadUserReputacion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LoadUserAntiFrags(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)
        '<EhHeader>
        On Error GoTo LoadUserAntiFrags_Err
        '</EhHeader>

        Dim A    As Long

        Dim Temp As String
    
100     With UserList(UserIndex).AntiFrags(A)
    
102         For A = 1 To Declaraciones.MAX_CONTROL_FRAGS
104             Temp = Userfile.GetValue("ANTIFRAGS", "FRAG" & A)
            
106             .UserName = ReadField(1, Temp, Asc("-"))
108             .Time = val(ReadField(2, Temp, Asc("-")))
110             .cant = val(ReadField(3, Temp, Asc("-")))
112             .Account = val(ReadField(4, Temp, Asc("-")))
114         Next A
        
        End With

        '<EhFooter>
        Exit Sub

LoadUserAntiFrags_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadUserAntiFrags " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SaveUserAntiFrags(ByRef AntiFrags() As tAntiFrags, _
                             ByRef Manager As clsIniManager)

        '<EhHeader>
        On Error GoTo SaveUserAntiFrags_Err

        '</EhHeader>

        Dim A    As Long

        Dim Temp As String
    
100
    
102     For A = 1 To Declaraciones.MAX_CONTROL_FRAGS

            With AntiFrags(A)
104             Call Manager.ChangeValue("ANTIFRAGS", "FRAG" & A, .UserName & "-" & .Time & "-" & .cant & "-" & .Account)

            End With

106     Next A

        '<EhFooter>
        Exit Sub

SaveUserAntiFrags_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.ES.SaveUserAntiFrags " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)

        '<EhHeader>
        On Error GoTo LoadUserInit_Err

        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 19/11/2006
        'Loads the Users RECORDs
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
        '*************************************************
        Dim LoopC As Long

        Dim ln    As String
    
        Dim A As Long
        
100     With UserList(UserIndex)
        
102         With .Stats
104             .Points = CLng(Userfile.GetValue("STATS", "Points"))
        
            End With
            
            With .Skins
                .Last = CByte(Userfile.GetValue("SKINS", "LAST"))
                .ArmourIndex = CInt(Userfile.GetValue("SKINS", "ARMOUR"))
                .ShieldIndex = CInt(Userfile.GetValue("SKINS", "SHIELD"))
                .WeaponIndex = CInt(Userfile.GetValue("SKINS", "WEAPON"))
                .WeaponArcoIndex = CInt(Userfile.GetValue("SKINS", "WEAPONARCO"))
                .WeaponDagaIndex = CInt(Userfile.GetValue("SKINS", "WEAPONDAGA"))
                .HelmIndex = CInt(Userfile.GetValue("SKINS", "HELM"))
                
                ReDim .ObjIndex(1 To MAX_INVENTORY_SKINS) As Integer
                    
                For LoopC = 1 To MAX_INVENTORY_SKINS
                    .ObjIndex(LoopC) = val(Userfile.GetValue("SKINS", CStr(LoopC)))
                Next LoopC

            End With
        
106         .RankMonth = CByte(Userfile.GetValue("RANKING", "RANKMONTH"))
        
108         With .Faction
110             .FragsCiu = CLng(Userfile.GetValue("FACTION", "FragsCiu"))
112             .FragsCri = CLng(Userfile.GetValue("FACTION", "FragsCri"))
114             .FragsOther = CLng(Userfile.GetValue("FACTION", "FragsOther"))
116             .Range = CByte(Userfile.GetValue("FACTION", "Range"))
118             .Status = CByte(Userfile.GetValue("FACTION", "Status"))
120             .StartDate = CStr(Userfile.GetValue("FACTION", "StartDate"))
122             .StartElv = CByte(Userfile.GetValue("FACTION", "StartElv"))
124             .StartFrags = CInt(Userfile.GetValue("FACTION", "StartFrags"))
126             .ExFaction = CByte(Userfile.GetValue("FACTION", "ExFaction"))

            End With

128         With .flags
                .Rachas = CInt(Userfile.GetValue("FLAGS", "Rachas"))
                .RachasTemp = CInt(Userfile.GetValue("FLAGS", "RachasTemp"))
130             .StreamUrl = Userfile.GetValue("FLAGS", "StreamUrl")
            
132             .Blocked = CByte(Userfile.GetValue("FLAGS", "BLOCKED"))
134             .ObjIndex = CInt(Userfile.GetValue("FLAGS", "ObjIndex"))
136             .SelectedBono = CInt(Userfile.GetValue("FLAGS", "SelectedBono"))
138             .Muerto = CByte(Userfile.GetValue("FLAGS", "Muerto"))
140             .Escondido = CByte(Userfile.GetValue("FLAGS", "Escondido"))
            
142             .Hambre = CByte(Userfile.GetValue("FLAGS", "Hambre"))
144             .Sed = CByte(Userfile.GetValue("FLAGS", "Sed"))
146             .Desnudo = CByte(Userfile.GetValue("FLAGS", "Desnudo"))
148             .Navegando = CByte(Userfile.GetValue("FLAGS", "Navegando"))
150             .Montando = CInt(Userfile.GetValue("FLAGS", "Montando"))
152             .Envenenado = CByte(Userfile.GetValue("FLAGS", "Envenenado"))
154             .Paralizado = CByte(Userfile.GetValue("FLAGS", "Paralizado"))
            
                'Matrix
156             .LastMap = val(Userfile.GetValue("FLAGS", "LastMap"))
            
158             .Streamer = CByte(Userfile.GetValue("FLAGS", "STREAMER"))
160             .Premium = CByte(Userfile.GetValue("FLAGS", "PREMIUM"))
162             .Oro = CByte(Userfile.GetValue("FLAGS", "ORO"))
164             .Bronce = CByte(Userfile.GetValue("FLAGS", "BRONCE"))
166             .Plata = CByte(Userfile.GetValue("FLAGS", "PLATA"))

            End With
        
168         If .flags.Paralizado = 1 Then
170             .Counters.Paralisis = IntervaloParalizado

            End If
        
            ' .Counters.Incinerado = CInt(UserFile.GetValue("COUNTERS", "INCINERADO"))
172         .Counters.TimePublicationMao = CInt(Userfile.GetValue("COUNTERS", "TIMEPUBLICATIONMAO"))
174         .Counters.TimeBonus = CLng(Userfile.GetValue("COUNTERS", "TIMEBONUS"))
176         .Counters.TimeTransform = CInt(Userfile.GetValue("COUNTERS", "TIMETRANSFORM"))
178         .Counters.TimeBono = CInt(Userfile.GetValue("COUNTERS", "TIMEBONO"))
180         .Counters.TimeTelep = CInt(Userfile.GetValue("COUNTERS", "TIMETELEP"))
182         .Counters.Pena = CLng(Userfile.GetValue("COUNTERS", "Pena"))
184         .Counters.AsignedSkills = CByte(val(Userfile.GetValue("COUNTERS", "SkillsAsignados")))
        
186         .Genero = Userfile.GetValue("INIT", "Genero")
188         .Clase = Userfile.GetValue("INIT", "Clase")
190         .Raza = Userfile.GetValue("INIT", "Raza")
192         .Hogar = Userfile.GetValue("INIT", "Hogar")
194         .Char.Heading = CInt(Userfile.GetValue("INIT", "Heading"))
        
196         With .OrigChar
198             .Head = CInt(Userfile.GetValue("INIT", "Head"))
200             .Body = CInt(Userfile.GetValue("INIT", "Body"))
202             .WeaponAnim = CInt(Userfile.GetValue("INIT", "Arma"))
204             .ShieldAnim = CInt(Userfile.GetValue("INIT", "Escudo"))
206             .CascoAnim = CInt(Userfile.GetValue("INIT", "Casco"))
            
208             .Heading = eHeading.SOUTH

            End With
        
            #If ConUpTime Then
210             .UpTime = CLng(Userfile.GetValue("INIT", "UpTime"))
            #End If
        
212         If .flags.Muerto = 0 Then
214             .Char = .OrigChar
            Else
216             .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
218             .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
            
222             .Char.WeaponAnim = NingunArma
224             .Char.ShieldAnim = NingunEscudo
226             .Char.CascoAnim = NingunCasco

                  For A = 1 To MAX_AURAS
228             .Char.AuraIndex(A) = NingunAura
                  Next A
            End If
        
230         .Desc = Userfile.GetValue("INIT", "Desc")
        
232         .Pos.Map = CInt(ReadField(1, Userfile.GetValue("INIT", "Position"), 45))
234         .Pos.X = CInt(ReadField(2, Userfile.GetValue("INIT", "Position"), 45))
236         .Pos.Y = CInt(ReadField(3, Userfile.GetValue("INIT", "Position"), 45))
        
238         .Invent.NroItems = CInt(Userfile.GetValue("Inventory", "CantidadItems"))
        
            '[KEVIN]--------------------------------------------------------------------
            '***********************************************************************************
240         .BancoInvent.NroItems = CInt(Userfile.GetValue("BancoInventory", "CantidadItems"))

            'Lista de objetos del banco
242         For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
244             ln = Userfile.GetValue("BancoInventory", "Obj" & LoopC)
246             .BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
248             .BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
250         Next LoopC

            '------------------------------------------------------------------------------------
            '[/KEVIN]*****************************************************************************
        
            'Lista de objetos
252         For LoopC = 1 To MAX_INVENTORY_SLOTS
254             ln = Userfile.GetValue("Inventory", "Obj" & LoopC)
256             .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
258             .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
260             .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
262         Next LoopC
        
            'Obtiene el indice-objeto del arma
264         .Invent.WeaponEqpSlot = CByte(Userfile.GetValue("Inventory", "WeaponEqpSlot"))

266         If .Invent.WeaponEqpSlot > 0 Then
268             .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex

            End If
        
            'Obtiene el indice-objeto del aura
270         .Invent.AuraEqpSlot = CByte(Userfile.GetValue("Inventory", "AuraSlot"))

272         If .Invent.AuraEqpSlot > 0 Then
274             .Invent.AuraEqpObjIndex = .Invent.Object(.Invent.AuraEqpSlot).ObjIndex
            
276             .Char.AuraIndex(5) = ObjData(.Invent.AuraEqpObjIndex).AuraIndex

            End If
        
            'Obtiene el indice-objeto del armadura
278         .Invent.ArmourEqpSlot = CByte(Userfile.GetValue("Inventory", "ArmourEqpSlot"))

280         If .Invent.ArmourEqpSlot > 0 Then
282             .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
284             .flags.Desnudo = 0
            Else
286             .flags.Desnudo = 1

            End If
        
            'Obtiene el indice-objeto del escudo
288         .Invent.EscudoEqpSlot = CByte(Userfile.GetValue("Inventory", "EscudoEqpSlot"))

290         If .Invent.EscudoEqpSlot > 0 Then
292             .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex

            End If
        
            'Obtiene el indice-objeto del casco
294         .Invent.CascoEqpSlot = CByte(Userfile.GetValue("Inventory", "CascoEqpSlot"))

296         If .Invent.CascoEqpSlot > 0 Then
298             .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex

            End If
        
            'Obtiene el indice-objeto del Anillo Magico o Laud Magico
300         .Invent.MagicSlot = CByte(Userfile.GetValue("Inventory", "MagicSlot"))

302         If .Invent.MagicSlot > 0 Then
304             .Invent.MagicObjIndex = .Invent.Object(.Invent.MagicSlot).ObjIndex

            End If
        
            'Obtiene el indice-objeto barco
306         .Invent.BarcoSlot = CByte(Userfile.GetValue("Inventory", "BarcoSlot"))

308         If .Invent.BarcoSlot > 0 Then
310             .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex

            End If
        
            'Obtiene el indice-objeto municion
312         .Invent.MunicionEqpSlot = CByte(Userfile.GetValue("Inventory", "MunicionSlot"))

314         If .Invent.MunicionEqpSlot > 0 Then
316             .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex

            End If
        
            '[Alejo]
            'Obtiene el indice-objeto anilo
318         .Invent.AnilloEqpSlot = CByte(Userfile.GetValue("Inventory", "AnilloSlot"))

320         If .Invent.AnilloEqpSlot > 0 Then
322             .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex

            End If
        
324         .Invent.MochilaEqpSlot = val(Userfile.GetValue("Inventory", "MochilaSlot"))

326         If .Invent.MochilaEqpSlot > 0 Then
328             .Invent.MochilaEqpObjIndex = .Invent.Object(.Invent.MochilaEqpSlot).ObjIndex

            End If
        
330         .Invent.MonturaSlot = val(Userfile.GetValue("Inventory", "MonturaSlot"))

332         If .Invent.MonturaSlot > 0 Then
334             .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex

            End If
            
            .Invent.PendientePartySlot = val(Userfile.GetValue("Inventory", "PendientePartySlot"))

            If .Invent.PendientePartySlot > 0 Then
                .Invent.PendientePartyObjIndex = .Invent.Object(.Invent.PendientePartySlot).ObjIndex

            End If
        
336         ln = Userfile.GetValue("Guild", "GUILDINDEX")

338         If IsNumeric(ln) Then
340             .GuildIndex = CInt(ln)
            Else
342             .GuildIndex = 0

            End If
        
344         ln = Userfile.GetValue("Guild", "GUILDRANGE")

346         If IsNumeric(ln) Then
348             .GuildRange = CByte(ln)
            Else
350             .GuildRange = rNone

            End If
        
        End With

        '<EhFooter>
        Exit Sub

LoadUserInit_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.ES.LoadUserInit " & "at line " & Erl

        

        '</EhFooter>
End Sub

Function GetVar(ByVal File As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo GetVar_Err
        '</EhHeader>

        Dim sSpaces  As String ' This will hold the input that the program will retrieve

        Dim szReturn As String ' This will be the defaul value if the string is not found
      
100     szReturn = vbNullString
      
102     sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
104     GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
106     GetVar = RTrim$(sSpaces)
108     GetVar = Left$(GetVar, Len(GetVar) - 1)
  
        '<EhFooter>
        Exit Function

GetVar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.GetVar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub CargarBackUp()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CargarBackUp_Err
        '</EhHeader>

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
    
        Dim Map       As Integer

        Dim tFileName As String
    
        On Error GoTo Man
        
102     NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
104     Call ModAreas.Initialise(NumMaps)
        
106     frmCargando.cargar.Min = 0
108     frmCargando.cargar.max = NumMaps
110     frmCargando.cargar.Value = 0
        
112     MapPath = Maps_FilePath
        
114     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
116     ReDim MapInfo(1 To NumMaps) As MapInfo
118     ReDim MiniMap(1 To NumMaps) As tMinimap
    
120     For Map = 1 To NumMaps

122         If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
124             tFileName = Maps_FilePath & "WORLDBACKUP\Mapa" & Map
                
126             If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
128                 tFileName = Maps_FilePath & "Mapa" & Map
                End If

            Else
130             tFileName = Maps_FilePath & "Mapa" & Map
            End If
            
132         Call CargarMapa(Map, tFileName)
            'Call GrabarMapa(Map, Maps_FilePath & "WORLDBACKUP\Mapa" & Map)
134         frmCargando.cargar.Value = frmCargando.cargar.Value + 1
136         DoEvents
138     Next Map
    
        Exit Sub

Man:
140     MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
142     Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
        '<EhFooter>
        Exit Sub

CargarBackUp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.CargarBackUp " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LoadMapData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
    
    Dim Map       As Integer

    Dim tFileName As String
    
    On Error GoTo Man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call ModAreas.Initialise(NumMaps)
        
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = Maps_FilePath
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
    'Call WriteAddmapConfig
    
    For Map = 1 To NumMaps
            
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
    
    'Call WriteAddObj_Finish
    
    Exit Sub

Man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByRef MAPFl As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/08/2010
    '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
    '***************************************************

    On Error GoTo errh

    Dim hFile     As Integer
    Dim X         As Long
    Dim Y         As Long
    Dim ByFlags   As Byte
    Dim npcfile   As String
    Dim Leer      As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff()    As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    npcfile = Npcs_FilePath
    
    Dim PosType As Byte 'Si la posicion es de tierra (0) o agua (1).
    ReDim MapInfo(Map).NpcSpawnPos(0).Pos(0)
    ReDim MapInfo(Map).NpcSpawnPos(1).Pos(0)
    
    hFile = FreeFile

    If Not FileExist(MAPFl & ".map", vbArchive) Then Exit Sub
    
    Open MAPFl & ".map" For Binary As #hFile
    Seek hFile, 1
    ReDim Buff(LOF(hFile) - 1) As Byte
    Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'infdleGuilds_Kick
    Open MAPFl & ".inf" For Binary As #hFile
    Seek hFile, 1
    ReDim Buff(LOF(hFile) - 1) As Byte
    Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(Map).MapVersion = MapReader.getInteger
    
    MiCabecera.Desc = MapReader.getString(Len(MiCabecera.Desc))
    MiCabecera.CRC = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    Dim B As Long
    
    Call Leer.Initialize(MAPFl & ".dat")
    
    With MapInfo(Map)
        .Name = Leer.GetValue("Mapa" & Map, "Name")
        .Music = Leer.GetValue("Mapa" & Map, "MusicNum")
        .StartPos.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        
        .OnDeathGoTo.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        
        .OnLoginGoTo.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "OnLoginGoTo"), Asc("-")))
        .OnLoginGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "OnLoginGoTo"), Asc("-")))
        .OnLoginGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "OnLoginGoTo"), Asc("-")))
        
        .GoToOns.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "GoToOns"), Asc("-")))
        .GoToOns.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "GoToOns"), Asc("-")))
        .GoToOns.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "GoToOns"), Asc("-")))
        
        .MagiaSinEfecto = val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
        .InviSinEfecto = val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
        .ResuSinEfecto = val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
        .OcultarSinEfecto = val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))
        .InvocarSinEfecto = val(Leer.GetValue("Mapa" & Map, "InvocarSinEfecto"))
        .MimetismoSinEfecto = val(Leer.GetValue("Mapa" & Map, "MimetismoSinEfecto"))
        .Faction = val(Leer.GetValue("Mapa" & Map, "Faction"))
        .RoboNpcsPermitido = val(Leer.GetValue("Mapa" & Map, "RoboNpcsPermitido"))
        .LvlMin = val(Leer.GetValue("Mapa" & Map, "LvlMin"))
        .LvlMax = val(Leer.GetValue("Mapa" & Map, "LvlMax"))
        .Guild = val(Leer.GetValue("Mapa" & Map, "Guild"))

        .SubMaps = val(Leer.GetValue("Mapa" & Map, "SUB_MAPS"))
        
        If .SubMaps > 0 Then
            Dim ArraiMaps() As String
            ArraiMaps = Split(Leer.GetValue("Mapa" & Map, "MAPS"), "-")
            
            ReDim .Maps(1 To .SubMaps) As Integer
            
            For B = 1 To .SubMaps
                .Maps(B) = val(ArraiMaps(B - 1))
            Next B
        End If
        
        .Pesca = val(Leer.GetValue("Mapa" & Map, "Pesca"))
        
        If .Pesca > 0 Then
            ReDim .PescaItem(1 To .Pesca) As Integer
            
            For B = 1 To .Pesca
                .PescaItem(B) = val(Leer.GetValue("Mapa" & Map, "P" & B))
            Next B
        End If
        
        'Call WriteVar(MAPFl & ".dat", "Mapa" & Map, "Limpieza", "1")
        ' Call WriteVar(MAPFl & ".dat", "Mapa" & Map, "CaenItems", "1")
        
        .Limpieza = val(Leer.GetValue("Mapa" & Map, "Limpieza"))
        .CaenItems = val(Leer.GetValue("Mapa" & Map, "CaenItems"))
        
        .Bronce = val(Leer.GetValue("Mapa" & Map, "Bronce"))
        .Plata = val(Leer.GetValue("Mapa" & Map, "Plata"))
        .Premium = val(Leer.GetValue("Mapa" & Map, "Premium"))
        
        If val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False
        End If
        
        .Terreno = TerrainStringToByte(Leer.GetValue("Mapa" & Map, "Terreno"))
        .Zona = Leer.GetValue("Mapa" & Map, "Zona")
        .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & Map, "Restringir"))
        .BackUp = val(Leer.GetValue("Mapa" & Map, "BACKUP"))
        .NoMana = val(Leer.GetValue("Mapa" & Map, "NOMANA"))
          
        .MinOns = val(Leer.GetValue("Mapa" & Map, "MINONLINES"))
        
        .Poder = val(Leer.GetValue("Mapa" & Map, "PODER"))
        
        Dim Days() As String
        Dim starts() As String
        Dim ends() As String
        Dim i As Integer

        Days = Split(Leer.GetValue("Mapa" & Map, "AccessDays"), "-")
        starts = Split(Leer.GetValue("Mapa" & Map, "AccessTimeStarts"), "-")
        ends = Split(Leer.GetValue("Mapa" & Map, "AccessTimeEnds"), "-")
    
        If UBound(Days) <> -1 Then
            ReDim .AccessDays(UBound(Days))
            ReDim .AccessTimeStarts(UBound(starts))
            ReDim .accessTimeEnds(UBound(ends))
    
            For i = LBound(Days) To UBound(Days)
                .AccessDays(i) = val(Days(i))
                .AccessTimeStarts(i) = val(starts(i))
                .accessTimeEnds(i) = val(ends(i))
            Next i
        Else
            ReDim .AccessDays(0)
            ReDim .AccessTimeStarts(0)
            ReDim .accessTimeEnds(0)
    
        End If
        
        Call MiniMap_SetInfo(Map)
    End With
    
    Set MapInfo(Map).Players = New Network.Group
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = MapReader.getLong

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getLong

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getLong

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getLong

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.Map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger
                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                    .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then
                        .NpcIndex = OpenNPC(.NpcIndex, LeerNPCs)
                        
                        If .NpcIndex > 10000 Then
                            .NpcIndex = OpenNPC(1, LeerNPCs)   ' buscarlas en el mapa cornelios
                        End If
                         
                        Npclist(.NpcIndex).Orig.Map = Map
                        Npclist(.NpcIndex).Orig.X = X
                        Npclist(.NpcIndex).Orig.Y = Y
                        
                        Npclist(.NpcIndex).Pos.Map = Map
                        Npclist(.NpcIndex).Pos.X = X
                        Npclist(.NpcIndex).Pos.Y = Y
                        
                        Call UpdateInfoNpcs(Map, .NpcIndex)
                        Call MakeNPCChar(True, 0, .NpcIndex, Map, X, Y)
                        
                    End If
                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.Amount = InfReader.getInteger
                    
                    If .ObjInfo.ObjIndex > 0 Then
                        If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otcofre Then
                            Call MiniMap_SetChest(Map, .ObjInfo.ObjIndex)
                        End If
                    End If
                    
                    ' TODO: Item and Object separation
                    Dim Coordinates As WorldPos

                    Coordinates.Map = Map
                    Coordinates.X = X
                    Coordinates.Y = Y
                    Call ModAreas.CreateEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT, Coordinates, ObjData(.ObjInfo.ObjIndex).SizeWidth, ObjData(.ObjInfo.ObjIndex).SizeHeight)
                End If
                     
                'Se fija si la posicion es valida para un npc de agua o tierra y la guarda por separado.
                PosType = LegalNpcSpawnPos(Map, X, Y)
                If PosType > 0 Then
                    ReDim Preserve MapInfo(Map).NpcSpawnPos(PosType - 1).Pos(0 To UBound(MapInfo(Map).NpcSpawnPos(PosType - 1).Pos) + 1)
                    MapInfo(Map).NpcSpawnPos(PosType - 1).Pos(UBound(MapInfo(Map).NpcSpawnPos(PosType - 1).Pos)).X = X
                    MapInfo(Map).NpcSpawnPos(PosType - 1).Pos(UBound(MapInfo(Map).NpcSpawnPos(PosType - 1).Pos)).Y = Y
                End If

            End With

        Next X
    Next Y
    
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff

    Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
End Sub

Public Function LegalNpcSpawnPos(ByVal Map As Long, ByVal X As Byte, ByVal Y As Byte) As Byte
    '***************************************************
    'Author: Anagrama
    'Last Modification: 07/01/2017
    'Revisa si la posición es válida para el spawn de un npc de tierra o agua de forma diferenciada.
    'Luego devuelve si es válida, si es de tierra o agua.
    '***************************************************
    On Error GoTo ErrHandler
    Dim IsLegal As Boolean
    With MapData(Map, X, Y)
        IsLegal = LegalPos(Map, X, Y, False, True, True)
        IsLegal = IsLegal And (.trigger <> eTrigger.POSINVALIDA)
        IsLegal = IsLegal And InMapBounds(Map, X, Y)
        If IsLegal = True Then
            LegalNpcSpawnPos = 1
            Exit Function
        End If
        IsLegal = LegalPos(Map, X, Y, True, False, True)
        IsLegal = IsLegal And (.trigger <> eTrigger.POSINVALIDA)
        IsLegal = IsLegal And InMapBounds(Map, X, Y)
        If IsLegal = True Then
            LegalNpcSpawnPos = 2
            Exit Function
        End If
    End With
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.number & "(" & Err.description & ") en Function LegalNpcSpawnPos de FileIO")
End Function

Sub LoadSini()
        '<EhHeader>
        On Error GoTo LoadSini_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim Temporal As Long
    
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
    
        lastRunTime = GetVar(IniPath & "Server.ini", "INIT", "lastRunTime")
        DateAperture = GetVar(IniPath & "Server.ini", "INIT", "FechaApertura")
102     BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))
    
104     TOLERANCE_MS_POTION = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TOLERANCE_MS_POTION"))
106     TOLERANCE_AMOUNT_POTION = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TOLERANCE_AMOUNT_POTION"))
     
108     TOLERANCE_POTIONBLUE_CLIC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TOLERANCE_BLUE_CLIC"))
110     TOLERANCE_POTIONBLUE_U = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TOLERANCE_BLUE_U"))

112     TOLERANCE_POTIONRED_CLIC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TOLERANCE_RED_CLIC"))
114     TOLERANCE_POTIONRED_U = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TOLERANCE_RED_U"))

116     ModoNavidad = val(GetVar(IniPath & "Server.ini", "FESTIVIDADES", "ModoNavidad"))
    
118     MultExp = val(GetVar(IniPath & "Server.ini", "INIT", "MultExp"))
120     MultGld = val(GetVar(IniPath & "Server.ini", "INIT", "MultOro"))
    
122     Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
        'Puerto = Open_Server_Port
        
124     HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
126     AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
128     IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
        'Lee la version correcta del cliente
130     ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
    
132     PuedeConectarPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeConectarPersonajes"))
134     PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
136     ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))
138     ValidacionDePjs = val(GetVar(IniPath & "Server.ini", "INIT", "ValidacionDePjs"))
    
140     With ConfigServer
142         .ModoRetos = val(GetVar(IniPath & "Server.ini", "CONFIG", "MODORETOS"))
144         .ModoRetosFast = val(GetVar(IniPath & "Server.ini", "CONFIG", "MODORETOSFAST"))
146         .ModoInvocaciones = val(GetVar(IniPath & "Server.ini", "CONFIG", "MODOINVOCACIONES"))
148         .ModoCastillo = val(GetVar(IniPath & "server.ini", "CONFIG", "MODOCASTILLO"))
150         .ModoCrafting = val(GetVar(IniPath & "Server.ini", "CONFIG", "MODOCRAFTING"))
152         .ModoSubastas = val(GetVar(IniPath & "Server.ini", "CONFIG", "MODOSUBASTAS"))
            .ModoSkins = val(GetVar(IniPath & "Server.ini", "CONFIG", "MODOSKINS"))
            Events_Automatic.Events_Automatic_Active = val(GetVar(IniPath & "Server.ini", "CONFIG", "JARVIS"))
              
            frmServidor.chkEvents.Value = Events_Automatic.Events_Automatic_Active
154         frmServidor.chkRetos.Value = .ModoRetos
156         frmServidor.chkRetosFast.Value = .ModoRetosFast
158         frmServidor.chkInvocaciones.Value = .ModoInvocaciones
160         frmServidor.chkCastillo.Value = .ModoCastillo
162         frmServidor.chkCrafting.Value = .ModoCrafting
164         frmServidor.chkSubastas.Value = .ModoSubastas
            frmServidor.chkSkins.Value = .ModoSkins

        End With
    
166     EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
    
        'Intervalos
168     SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
170     FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
172     StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
174     FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
176     SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
178     FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
180     StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
182     FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
184     IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
186     FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
188     IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
190     FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
192     IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
194     FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
196     IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
198     FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
200     IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
202     FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
204     IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
206     FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
208     IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
210     FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
212     IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
214     FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
216     IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
218     FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion
    
        '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
220     IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
222     IntervaloAtacable = 60000 ' Cargar desde balance.dat
224     IntervaloOwnedNpc = 18000 ' Cargar desde balance.dat
    
        IntervaloUserPuedeShiftear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPuedeShiftear"))

226     IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
228     FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
230     IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
232     FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
234     IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
236     FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
238     IntervalDrop = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloDrop"))
240     FrmInterv.txtDrop.Text = IntervalDrop
    
242     IntervaloPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPuedeCastear"))
244     FrmInterv.txtCast.Text = IntervaloPuedeCastear
         
         IntervaloMeditar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMeditar"))
         IntervaloCaminar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCaminar"))
         MaximoSpeedHack = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "MaximoSpeedHack"))
         
246     IntervalCommerce = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloComercio"))
248     FrmInterv.txtCommerce.Text = IntervalCommerce
    
250     IntervalMessage = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloChat"))
252     FrmInterv.txtMessage.Text = IntervalMessage
    
254     IntervalInfoMao = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInfoMao"))
256     FrmInterv.txtInfoMao.Text = IntervalInfoMao
    
          IntervaloEquipped = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloEquipped"))
          
        'TODO : Agregar estos intervalos al form!!!
258     IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
260     IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
262     IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))
264     MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
    
266     IntervaloGuardarUsuarios = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGuardarUsuarios"))
268     IntervaloTimerGuardarUsuarios = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerGuardarUsuarios"))
    
270     IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
272     IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
274     IntervaloUserPuedeUsarClick = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsarClick"))
276     IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))
    
278     IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
280     RECORDusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "RECORD"))
      
        'Max users
282     Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))

284     If MaxUsers = 0 Then
286         MaxUsers = Temporal
288         ReDim UserList(1 To MaxUsers) As User
290         ReDim AccountList(1 To MaxUsers) As tAccount

        End If
    
292     Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
294     Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
296     Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
298     Nix.Map = GetVar(DatPath & "Ciudades.dat", "Nix", "Mapa")
300     Nix.X = GetVar(DatPath & "Ciudades.dat", "Nix", "X")
302     Nix.Y = GetVar(DatPath & "Ciudades.dat", "Nix", "Y")
    
304     Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
306     Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
308     Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
310     Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
312     Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
314     Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
316     Arghal.Map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
318     Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
320     Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")

322     Esperanza.Map = GetVar(DatPath & "Ciudades.dat", "Esperanza", "Mapa")
324     Esperanza.X = GetVar(DatPath & "Ciudades.dat", "Esperanza", "X")
326     Esperanza.Y = GetVar(DatPath & "Ciudades.dat", "Esperanza", "Y")


328     Newbie.Map = GetVar(DatPath & "Ciudades.dat", "Newbie", "Mapa")
330     Newbie.X = GetVar(DatPath & "Ciudades.dat", "Newbie", "X")
332     Newbie.Y = GetVar(DatPath & "Ciudades.dat", "Newbie", "Y")

        
        CiudadFlotante.Map = GetVar(DatPath & "Ciudades.dat", "CiudadFlotante", "Mapa")
        CiudadFlotante.X = GetVar(DatPath & "Ciudades.dat", "CiudadFlotante", "X")
        CiudadFlotante.Y = GetVar(DatPath & "Ciudades.dat", "CiudadFlotante", "Y")
        
340     Ciudades(eCiudad.cUllathorpe) = Ullathorpe
342     Ciudades(eCiudad.cNix) = Nix
344     Ciudades(eCiudad.cBanderbill) = Banderbill
346     Ciudades(eCiudad.cLindos) = Lindos
348     Ciudades(eCiudad.cArghal) = Arghal
350     Ciudades(eCiudad.cArkhein) = Arkhein
352     Ciudades(eCiudad.cNewbie) = Newbie
354     Ciudades(eCiudad.cEsperanza) = Esperanza
356     Call LoadElu
    
        ' Admins
358     Call loadAdministrativeUsers

        '<EhFooter>
        Exit Sub

LoadSini_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadSini " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        'Escribe VAR en un archivo
        '***************************************************
        '<EhHeader>
        On Error GoTo WriteVar_Err
        '</EhHeader>

100     writeprivateprofilestring Main, Var, Value, File
    
        '<EhFooter>
        Exit Sub

WriteVar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.WriteVar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SaveUser(ByRef IUser As User, _
             ByVal Userfile As String, _
             Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Unknown
    'Last modified: 10/10/2010 (Pato)
    'Saves the Users RECORDs
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '11/19/2009: Pato - Save the EluSkills and ExpSkills
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '10/10/2010: Pato - Saco el WriteVar e implemento la clase clsIniManager
    '*************************************************

    On Error GoTo ErrHandler
    
    Dim Manager As clsIniManager

    Dim Existe  As Boolean
    
    Dim A As Long
    
    With IUser

        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
        If .Clase = 0 Or .Stats.Elv = 0 Then
            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)

            Exit Sub

        End If
    
        Set Manager = New clsIniManager
    
        If FileExist(Userfile) Then
            Call Manager.Initialize(Userfile)
        
            If FileExist(Userfile & ".bk") Then Call Kill(Userfile & ".bk")
            Name Userfile As Userfile & ".bk"
        
            Existe = True

        End If
    
        If .flags.Mimetizado = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False

        End If
    
        Dim LoopC As Integer
    
        With .Skins
            Call Manager.ChangeValue("SKINS", "LAST", CStr(.Last))
            Call Manager.ChangeValue("SKINS", "ARMOUR", CStr(.ArmourIndex))
            Call Manager.ChangeValue("SKINS", "SHIELD", CStr(.ShieldIndex))
            Call Manager.ChangeValue("SKINS", "WEAPON", CStr(.WeaponIndex))
            Call Manager.ChangeValue("SKINS", "WEAPONDAGA", CStr(.WeaponDagaIndex))
            Call Manager.ChangeValue("SKINS", "WEAPONARCO", CStr(.WeaponArcoIndex))
            Call Manager.ChangeValue("SKINS", "HELM", CStr(.HelmIndex))
            
            
            For LoopC = 1 To MAX_INVENTORY_SKINS
                Call Manager.ChangeValue("SKINS", CStr(LoopC), CStr(.ObjIndex(LoopC)))
            Next LoopC

        End With
        
        Call Manager.ChangeValue("STATS", "POINTS", CStr(.Stats.Points))
        Call Manager.ChangeValue("RANKING", "RANKMONTH", CStr(.RankMonth))
        
        Call Manager.ChangeValue("FLAGS", "BLOCKED", CStr(.flags.Blocked))
        Call Manager.ChangeValue("FLAGS", "OBJINDEX", CStr(.flags.ObjIndex))
        Call Manager.ChangeValue("FLAGS", "SelectedBono", CStr(.flags.SelectedBono))
        Call Manager.ChangeValue("FLAGS", "ORO", CStr(.flags.Oro))
        Call Manager.ChangeValue("FLAGS", "PREMIUM", CStr(.flags.Premium))
        Call Manager.ChangeValue("FLAGS", "STREAMER", CStr(.flags.Streamer))
        Call Manager.ChangeValue("FLAGS", "BRONCE", CStr(.flags.Bronce))
        Call Manager.ChangeValue("FLAGS", "PLATA", CStr(.flags.Plata))
        Call Manager.ChangeValue("FLAGS", "Muerto", CStr(.flags.Muerto))
        Call Manager.ChangeValue("FLAGS", "Escondido", CStr(.flags.Escondido))
        Call Manager.ChangeValue("FLAGS", "Hambre", CStr(.flags.Hambre))
        Call Manager.ChangeValue("FLAGS", "Sed", CStr(.flags.Sed))
        Call Manager.ChangeValue("FLAGS", "Desnudo", CStr(.flags.Desnudo))
        Call Manager.ChangeValue("FLAGS", "Ban", CStr(.flags.Ban))
        Call Manager.ChangeValue("FLAGS", "Navegando", CStr(.flags.Navegando))
        Call Manager.ChangeValue("FLAGS", "Montando", CStr(.flags.Montando))
        Call Manager.ChangeValue("FLAGS", "Envenenado", CStr(.flags.Envenenado))
        Call Manager.ChangeValue("FLAGS", "Paralizado", CStr(.flags.Paralizado))
        Call Manager.ChangeValue("FLAGS", "StreamUrl", .flags.StreamUrl)
        Call Manager.ChangeValue("FLAGS", "Rachas", CStr(.flags.Rachas))
        Call Manager.ChangeValue("FLAGS", "RachasTemp", CStr(.flags.RachasTemp))
        
        'Matrix
        Call Manager.ChangeValue("FLAGS", "LastMap", CStr(.flags.LastMap))
    
        Call Manager.ChangeValue("CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
        Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
        
        Call Manager.ChangeValue("COUNTERS", "Incinerado", CStr(.Counters.Incinerado))
        Call Manager.ChangeValue("COUNTERS", "TIMEPUBLICATIONMAO", CStr(.Counters.TimePublicationMao))
        Call Manager.ChangeValue("COUNTERS", "TIMEBONUS", CStr(.Counters.TimeBonus))
        Call Manager.ChangeValue("COUNTERS", "TIMETRANSFORM", CStr(.Counters.TimeTransform))
        Call Manager.ChangeValue("COUNTERS", "TIMEBONO", CStr(.Counters.TimeBono))
        Call Manager.ChangeValue("COUNTERS", "TIMETELEP", CStr(.Counters.TimeTelep))
        Call Manager.ChangeValue("COUNTERS", "Pena", CStr(.Counters.Pena))
        Call Manager.ChangeValue("COUNTERS", "SkillsAsignados", CStr(.Counters.AsignedSkills))
        
        Call Manager.ChangeValue("FACTION", "FragsCiu", CStr(.Faction.FragsCiu))
        Call Manager.ChangeValue("FACTION", "FragsCri", CStr(.Faction.FragsCri))
        Call Manager.ChangeValue("FACTION", "FragsOther", CStr(.Faction.FragsOther))
        Call Manager.ChangeValue("FACTION", "Range", CStr(.Faction.Range))
        Call Manager.ChangeValue("FACTION", "Status", CStr(.Faction.Status))
        Call Manager.ChangeValue("FACTION", "StartFrags", CStr(.Faction.StartFrags))
        Call Manager.ChangeValue("FACTION", "StartElv", CStr(.Faction.StartElv))
        Call Manager.ChangeValue("FACTION", "StartDate", CStr(.Faction.StartDate))
        Call Manager.ChangeValue("FACTION", "ExFaction", CStr(.Faction.ExFaction))
    
        '¿Fueron modificados los atributos del usuario?
        If Not .flags.TomoPocion Then

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
            Next LoopC

        Else

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
            Next LoopC

        End If
    
         For LoopC = 1 To UBound(.Stats.UserSkillsEspecial)
            Call Manager.ChangeValue("SKILLSESPECIAL", "SKESP" & LoopC, CStr(.Stats.UserSkillsEspecial(LoopC)))
        Next LoopC
        
        For LoopC = 1 To UBound(.Stats.UserSkills)
            Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
            Call Manager.ChangeValue("SKILLS", "ELUSK" & LoopC, CStr(.Stats.EluSkills(LoopC)))
            Call Manager.ChangeValue("SKILLS", "EXPSK" & LoopC, CStr(.Stats.ExpSkills(LoopC)))
        Next LoopC


        Call Manager.ChangeValue("STATS", "BONUSTIPE", .Stats.BonusTipe)
        Call Manager.ChangeValue("STATS", "BONUSVALUE", .Stats.BonusValue)
        
        Call Manager.ChangeValue("INIT", "Genero", .Genero)
        Call Manager.ChangeValue("INIT", "Raza", .Raza)
        Call Manager.ChangeValue("INIT", "Hogar", .Hogar)
        Call Manager.ChangeValue("INIT", "Clase", .Clase)
        Call Manager.ChangeValue("INIT", "Desc", .Desc)
    
        Call Manager.ChangeValue("INIT", "Heading", CStr(.Char.Heading))
        Call Manager.ChangeValue("INIT", "Head", CStr(.OrigChar.Head))
    
        If .flags.Muerto = 0 Then
            If .Char.Body <> 0 Then
                Call Manager.ChangeValue("INIT", "Body", CStr(.Char.Body))

            End If

        End If
    
        Call Manager.ChangeValue("INIT", "Arma", CStr(.Char.WeaponAnim))
        Call Manager.ChangeValue("INIT", "Escudo", CStr(.Char.ShieldAnim))
        Call Manager.ChangeValue("INIT", "Casco", CStr(.Char.CascoAnim))
    
        #If ConUpTime Then
    
            If SaveTimeOnline Then

                Dim TempDate As Date

                TempDate = Now - .LogOnTime
                .LogOnTime = Now
                .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
                .UpTime = .UpTime
                Call Manager.ChangeValue("INIT", "UpTime", .UpTime)

            End If

        #End If
    
        'First time around?
        If Manager.GetValue("INIT", "LastIP1") = vbNullString Then
            Call Manager.ChangeValue("INIT", "LastIP1", .IpAddress & " - " & Date & ":" & Time)
            'Is it a different ip from last time?
        ElseIf .IpAddress <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then

            Dim i As Integer

            For i = 5 To 2 Step -1
                Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
            Next i

            Call Manager.ChangeValue("INIT", "LastIP1", .IpAddress & " - " & Date & ":" & Time)
            'Same ip, just update the date
        Else
            Call Manager.ChangeValue("INIT", "LastIP1", .IpAddress & " - " & Date & ":" & Time)

        End If
    
        Call Manager.ChangeValue("INIT", "Position", .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
    
        'Call Manager.ChangeValue("RANKING", "RETOS1JUGADOS", CStr(.Stats.Retos1Jugados))
        'Call Manager.ChangeValue("RANKING", "RETOS1GANADOS", CStr(.Stats.Retos1Ganados))
        'Call Manager.ChangeValue("RANKING", "RETOS2JUGADOS", CStr(.Stats.DesafiosJugados))
        'Call Manager.ChangeValue("RANKING", "RETOS2GANADOS", CStr(.Stats.DesafiosGanados))
        'Call Manager.ChangeValue("RANKING", "TORNEOSJUGADOS", CStr(.Stats.TorneosJugados))
        'Call Manager.ChangeValue("RANKING", "TORNEOSGANADOS", CStr(.Stats.TorneosGanados))
    
        Call Manager.ChangeValue("STATS", "BONOSHP", CStr(.Stats.BonosHp))
        Call Manager.ChangeValue("STATS", "ELDHIR", CStr(.Stats.Eldhir))
        Call Manager.ChangeValue("STATS", "GLD", CStr(.Stats.Gld))
    
        Call Manager.ChangeValue("STATS", "MaxHP", CStr(.Stats.MaxHp))
        Call Manager.ChangeValue("STATS", "MinHP", CStr(.Stats.MinHp))
    
        Call Manager.ChangeValue("STATS", "MaxSTA", CStr(.Stats.MaxSta))
        Call Manager.ChangeValue("STATS", "MinSTA", CStr(.Stats.MinSta))
    
        Call Manager.ChangeValue("STATS", "MaxMAN", CStr(.Stats.MaxMan))
        Call Manager.ChangeValue("STATS", "MinMAN", CStr(.Stats.MinMan))
    
        Call Manager.ChangeValue("STATS", "MaxHIT", CStr(.Stats.MaxHit))
        Call Manager.ChangeValue("STATS", "MinHIT", CStr(.Stats.MinHit))
    
        Call Manager.ChangeValue("STATS", "MaxAGU", CStr(.Stats.MaxAGU))
        Call Manager.ChangeValue("STATS", "MinAGU", CStr(.Stats.MinAGU))
    
        Call Manager.ChangeValue("STATS", "MaxHAM", CStr(.Stats.MaxHam))
        Call Manager.ChangeValue("STATS", "MinHAM", CStr(.Stats.MinHam))
    
        Call Manager.ChangeValue("STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
      
        Call Manager.ChangeValue("STATS", "EXP", CStr(.Stats.Exp))
        Call Manager.ChangeValue("STATS", "ELV", CStr(.Stats.Elv))
    
        Call Manager.ChangeValue("BONUS", "BONUSLAST", CStr(.Stats.BonusLast))
        
        
        For A = 1 To .Stats.BonusLast
            Call Manager.ChangeValue("BONUS", "BONUS" & A, CStr(.Stats.Bonus(A).Tipo) & "|" & _
                                                           CStr(.Stats.Bonus(A).Value) & "|" & _
                                                           CStr(.Stats.Bonus(A).Amount) & "|" & _
                                                           CStr(.Stats.Bonus(A).DurationSeconds) & "|" & _
                                                           CStr(.Stats.Bonus(A).DurationDate))
        Next A
        
        Call Manager.ChangeValue("STATS", "ELU", CStr(.Stats.Elu))
        Call Manager.ChangeValue("MUERTES", "UserMuertes", CStr(.Faction.FragsOther))
        Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
      
        '[KEVIN]----------------------------------------------------------------------------
        '*******************************************************************************************
        Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))

        Dim loopd As Integer

        For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
            Call Manager.ChangeValue("BancoInventory", "Obj" & loopd, .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount)
        Next loopd

        '*******************************************************************************************
        '[/KEVIN]-----------
      
        'Save Inv
        Call Manager.ChangeValue("Inventory", "CantidadItems", val(.Invent.NroItems))
    
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Call Manager.ChangeValue("Inventory", "Obj" & LoopC, .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped)
        Next LoopC
    
        Call Manager.ChangeValue("Inventory", "AuraSlot", CStr(.Invent.AuraEqpSlot))
        Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
        Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
        Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
        Call Manager.ChangeValue("Inventory", "MagicSlot", CStr(.Invent.MagicSlot))
        Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
        Call Manager.ChangeValue("Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
        Call Manager.ChangeValue("Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
        Call Manager.ChangeValue("Inventory", "MochilaSlot", CStr(.Invent.MochilaEqpSlot))
        Call Manager.ChangeValue("Inventory", "MonturaSlot", CStr(.Invent.MonturaSlot))
        Call Manager.ChangeValue("Inventory", "ReliquiaSlot", CStr(.Invent.ReliquiaSlot))
        Call Manager.ChangeValue("Inventory", "PendientePartySlot", CStr(.Invent.PendientePartySlot))
        Call Manager.ChangeValue("Inventory", "AnilloSlot", CStr(.Invent.AnilloEqpSlot))
    
        'Reputacion
        Call Manager.ChangeValue("REP", "Asesino", CStr(.Reputacion.AsesinoRep))
        Call Manager.ChangeValue("REP", "Bandido", CStr(.Reputacion.BandidoRep))
        Call Manager.ChangeValue("REP", "Burguesia", CStr(.Reputacion.BurguesRep))
        Call Manager.ChangeValue("REP", "Ladrones", CStr(.Reputacion.LadronesRep))
        Call Manager.ChangeValue("REP", "Nobles", CStr(.Reputacion.NobleRep))
        Call Manager.ChangeValue("REP", "Plebe", CStr(.Reputacion.PlebeRep))
    
        ' Meditaciones
        Call Manager.ChangeValue("MEDITATION", "SELECTED", CStr(.MeditationSelected))
    
        For LoopC = 1 To MAX_MEDITATION
            Call Manager.ChangeValue("MEDITATION", LoopC, CStr(.MeditationUser(LoopC)))
        Next LoopC
    
        Dim L As Long

        L = (-.Reputacion.AsesinoRep) + (-.Reputacion.BandidoRep) + .Reputacion.BurguesRep + (-.Reputacion.LadronesRep) + .Reputacion.NobleRep + .Reputacion.PlebeRep
        L = L / 6
        Call Manager.ChangeValue("REP", "Promedio", CStr(L))
    
        Dim cad As String
    
        For LoopC = 1 To MAXUSERHECHIZOS
            cad = .Stats.UserHechizos(LoopC)
            Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
        Next
    
        ' Quests / Misiones
        Call SaveQuestStats(IUser.QuestStats, Manager)
    
        ' Anti Frags
        Call SaveUserAntiFrags(IUser.AntiFrags, Manager)
    
        'Guarda los mensajes privados del usuario.
        Call GuardarMensajes(IUser, Manager)
        
        .Counters.LastSave = GetTime

    End With

    Call Manager.DumpFile(Userfile)

    Set Manager = Nothing

    If Existe Then Call Kill(Userfile & ".bk")

    Exit Sub

ErrHandler:
    Call LogError("Error en SaveUser")
    Set Manager = Nothing

End Sub

Function Escriminal(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo criminal_Err
        '</EhHeader>

        Dim L As Long
    
100     With UserList(UserIndex).Reputacion
102         L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
104         L = L / 6
106         Escriminal = (L < 0) Or UserList(UserIndex).Faction.Status = r_Caos
        
        End With

        '<EhFooter>
        Exit Function

criminal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.criminal " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 10/09/2010
        '10/09/2010 - Pato: Optimice el BackUp de NPCs
        '***************************************************
        '<EhHeader>
        On Error GoTo BackUPnPc_Err
        '</EhHeader>

        Dim LoopC As Integer
    
100     Print #hFile, "[NPC" & Npclist(NpcIndex).numero & "]"
    
102     With Npclist(NpcIndex)
            'General
104         Print #hFile, "Name=" & .Name
106         Print #hFile, "Desc=" & .Desc
108         Print #hFile, "Head=" & val(.Char.Head)
110         Print #hFile, "Body=" & val(.Char.Body)
112         Print #hFile, "Heading=" & val(.Char.Heading)
114         Print #hFile, "Movement=" & val(.Movement)
116         Print #hFile, "Attackable=" & val(.Attackable)
118         Print #hFile, "Comercia=" & val(.Comercia)
120         Print #hFile, "TipoItems=" & val(.TipoItems)
122         Print #hFile, "Hostil=" & val(.Hostile)
124         Print #hFile, "GiveEXP=" & val(.GiveEXP)
126         Print #hFile, "GiveGLD=" & val(.GiveGLD)
128         Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
130         Print #hFile, "NpcType=" & val(.NPCtype)
        
            'Stats
132         Print #hFile, "Alineacion=" & val(.flags.AIAlineacion)
134         Print #hFile, "DEF=" & val(.Stats.def)
136         Print #hFile, "MaxHit=" & val(.Stats.MaxHit)
138         Print #hFile, "MaxHp=" & val(.Stats.MaxHp)
140         Print #hFile, "MinHit=" & val(.Stats.MinHit)
142         Print #hFile, "MinHp=" & val(.Stats.MinHp)
        
            'Flags
144         Print #hFile, "ReSpawn=" & val(.flags.Respawn)
146         Print #hFile, "BackUp=" & val(.flags.BackUp)
148         Print #hFile, "Domable=" & val(.flags.Domable)
        
            'Inventario
150         Print #hFile, "NroItems=" & val(.Invent.NroItems)

152         If .Invent.NroItems > 0 Then

154             For LoopC = 1 To .Invent.NroItems
156                 Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount
158             Next LoopC

            End If
        
160         Print #hFile, ""
        End With

        '<EhFooter>
        Exit Sub

BackUPnPc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.BackUPnPc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LogBan(ByVal BannedIndex As Integer, _
           ByVal UserIndex As Integer, _
           ByVal Motivo As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LogBan_Err
        '</EhHeader>

100     Call WriteVar(LogPath & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
102     Call WriteVar(LogPath & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)
    
        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open LogPath & "GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, UserList(BannedIndex).Name
110     Close #mifile

        '<EhFooter>
        Exit Sub

LogBan_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LogBan " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub LogBanFromName(ByVal BannedName As String, _
                   ByVal UserIndex As Integer, _
                   ByVal Motivo As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LogBanFromName_Err
        '</EhHeader>

100     Call WriteVar(LogPath & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
102     Call WriteVar(LogPath & "BanDetail.dat", BannedName, "Reason", Motivo)
    
        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open LogPath & "GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        '<EhFooter>
        Exit Sub

LogBanFromName_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LogBanFromName " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo Ban_Err
        '</EhHeader>

100     Call WriteVar(LogPath & "BanDetail.dat", BannedName, "BannedBy", Baneador)
102     Call WriteVar(LogPath & "BanDetail.dat", BannedName, "Reason", Motivo)
    
        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open LogPath & "GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        '<EhFooter>
        Exit Sub

Ban_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.Ban " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub CargaApuestas()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CargaApuestas_Err
        '</EhHeader>

100     Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
102     Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
104     Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

        '<EhFooter>
        Exit Sub

CargaApuestas_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.CargaApuestas " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub generateMatrix(ByVal mapa As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo generateMatrix_Err
        '</EhHeader>

        Dim i As Integer

        Dim j As Integer
    
100     ReDim distanceToCities(1 To NumMaps) As HomeDistance
    
102     For j = 1 To NUMCIUDADES
104         For i = 1 To NumMaps
106             distanceToCities(i).distanceToCity(j) = -1
108         Next i
110     Next j
    
112     For j = 1 To NUMCIUDADES
114         For i = 1 To 4

116             Select Case i

                    Case eHeading.NORTH
118                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.NORTH), j, i, 0, 1)

120                 Case eHeading.EAST
122                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.EAST), j, i, 1, 0)

124                 Case eHeading.SOUTH
126                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.SOUTH), j, i, 0, 1)

128                 Case eHeading.WEST
130                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.WEST), j, i, -1, 0)
                End Select

132         Next i
134     Next j

        '<EhFooter>
        Exit Sub

generateMatrix_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.generateMatrix " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub setDistance(ByVal mapa As Integer, _
                       ByVal city As Byte, _
                       ByVal side As Integer, _
                       Optional ByVal X As Integer = 0, _
                       Optional ByVal Y As Integer = 0)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo setDistance_Err
        '</EhHeader>

        Dim i   As Integer

        Dim lim As Integer

100     If mapa <= 0 Or mapa > NumMaps Then Exit Sub

102     If distanceToCities(mapa).distanceToCity(city) >= 0 Then Exit Sub

104     If mapa = Ciudades(city).Map Then
106         distanceToCities(mapa).distanceToCity(city) = 0
        Else
108         distanceToCities(mapa).distanceToCity(city) = Abs(X) + Abs(Y)
        End If

110     For i = 1 To 4
112         lim = getLimit(mapa, i)

114         If lim > 0 Then

116             Select Case i

                    Case eHeading.NORTH
118                     Call setDistance(lim, city, i, X, Y + 1)

120                 Case eHeading.EAST
122                     Call setDistance(lim, city, i, X + 1, Y)

124                 Case eHeading.SOUTH
126                     Call setDistance(lim, city, i, X, Y - 1)

128                 Case eHeading.WEST
130                     Call setDistance(lim, city, i, X - 1, Y)
                End Select

            End If

132     Next i

        '<EhFooter>
        Exit Sub

setDistance_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.setDistance " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function getLimit(ByVal mapa As Integer, ByVal side As Byte) As Integer
        '<EhHeader>
        On Error GoTo getLimit_Err
        '</EhHeader>

        '***************************************************
        'Author: Budi
        'Last Modification: 31/01/2010
        'Retrieves the limit in the given side in the given map.
        'TODO: This should be set in the .inf map file.
        '***************************************************
        Dim X As Long

        Dim Y As Long

100     If mapa <= 0 Then Exit Function

102     For X = 15 To 87
104         For Y = 0 To 3

106             Select Case side

                    Case eHeading.NORTH
108                     getLimit = MapData(mapa, X, 7 + Y).TileExit.Map

110                 Case eHeading.EAST
112                     getLimit = MapData(mapa, 92 - Y, X).TileExit.Map

114                 Case eHeading.SOUTH
116                     getLimit = MapData(mapa, X, 94 - Y).TileExit.Map

118                 Case eHeading.WEST
120                     getLimit = MapData(mapa, 9 + Y, X).TileExit.Map
                End Select

122             If getLimit > 0 Then Exit Function
124         Next Y
126     Next X

        '<EhFooter>
        Exit Function

getLimit_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.getLimit " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub LoadAnimations()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 11/06/2011
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LoadAnimations_Err
        '</EhHeader>
100     AnimHogar(eHeading.NORTH) = 40
102     AnimHogar(eHeading.EAST) = 42
104     AnimHogar(eHeading.SOUTH) = 39
106     AnimHogar(eHeading.WEST) = 41
    
108     AnimHogarNavegando(eHeading.NORTH) = 44
110     AnimHogarNavegando(eHeading.EAST) = 46
112     AnimHogarNavegando(eHeading.SOUTH) = 43
114     AnimHogarNavegando(eHeading.WEST) = 45
        '<EhFooter>
        Exit Sub

LoadAnimations_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.ES.LoadAnimations " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
