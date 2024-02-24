Attribute VB_Name = "mBonus"
Option Explicit

' # Sistema de bonus aleatoreos, cargados y preconfigurados desde un .ini

' # Bonus del personaje activo, por TIPO-DURATION o DURATIONDATE
' # Carga Bonus1=Tipo|Value|ObjIndex|Amount|DurationUse|DurationDate
Public Type UserBonus
    Tipo As eBonusType
    Value As Long
    Amount As Integer
    
    DurationSeconds As Long         ' Duración en segundos. Esto descuento por uso online del personaje.
    DurationDate As String      ' Duración en fecha (Si termina un dia determinado) (Formato 22-22-2222 18:00:00hs)
    
End Type

' Constantes para los límites de la duración
Const MinDuration As Integer = 900
Const MaxDuration As Integer = 18000


Public Type tObjsReward
    ObjIndex As Integer ' El índice del objeto en la lista de objetos
    Level As Integer ' Nivel de calidad del objeto
End Type


' Tipos de Bonus que puedo otorgar
Public Enum eBonusType
    eGld = 1        ' Altera el % de Oro
    eExp = 2        ' Altera el % de Exp
    ePoints = 3     ' Altera el % de Puntos
    eDsp = 4        ' Altera el % de Dsp
    eObj = 5        ' Determina si da un objeto y no un efecto.
    eVip = 6        ' Agrega tiempo V.I.P
    eMap = 7        ' Agrega acceso a un mapa específico.
End Enum

' Estructura del Bonus
Public Type tBonus
    Tipo As eBonusType      ' Tipo de bonificacion
    PorcInitial As Byte     ' Valor Inicial del que tiene que partir
    AddStep As Byte         ' Valor que agrega cada vez que loopea
    MaxStep As Byte         ' Cantidad de veces que loopea para mejorar el porcentaje
    
    Level As Byte           ' Calidad del Drop. Es usado para determinar la calidad del Drop PreConfig (En algunos casos requerimos Efectos más precarios, y en otros más jugosos)
    Porc As Long
    Duration As Integer     ' En Segundos. Cuanyo Mayor Step sea, menor duración tiene que otorgar (El intervalo tiene que ser entre 15 minutos para el full jugado y 2 a 5 horas para el máximo)
    
    ObjIndex As Integer
End Type

' Estructura para almacenar las configuraciones de nivel
Public Type BonusConfig
    MinPorcInitial As Integer
    MaxPorcInitial As Integer
    MaxAddStep As Integer
    MaxMaxStep As Integer
End Type

Public LevelConfig As BonusConfig
Public ObjsReward() As tObjsReward


' # Función para generar una configuración aleatoria de bonus
Public Function Bonus_GenerateRandomConfigs(ByVal Tipo As eBonusType, _
                                            ByVal Count As Integer, _
                                            ByVal Level As Integer, _
                                            ByVal Prob As Single) As tBonus()
    Dim A As Integer
    Dim Configs() As tBonus
    ReDim Configs(1 To Count) As tBonus
    
    Dim ProbThreshold As Single
    ProbThreshold = Prob * 1000 ' Ajustar el umbral de probabilidad (10% se convierte en 10000)
    
    For A = 1 To Count
        With Configs(A)
            If RandomNumber(1, 10000) <= ProbThreshold Then
                .Tipo = Tipo
                .Level = Level
                
                Select Case .Tipo
                     Case eBonusType.eGld, eBonusType.eExp, eBonusType.ePoints, eBonusType.eDsp
                        Dim stepChoices() As Variant
                        Dim curLevelConfig As BonusConfig
                        curLevelConfig = GetLevelConfig(Level)
                        stepChoices = Array(5, 10)    ' Seleccionar un incremento para Porc de entre estos valores
        
                        .PorcInitial = RandomNumber(curLevelConfig.MinPorcInitial, curLevelConfig.MaxPorcInitial)
                        .AddStep = stepChoices(RandomNumber(LBound(stepChoices), UBound(stepChoices)))
                        .MaxStep = RandomNumber(1, curLevelConfig.MaxMaxStep)
                        .Porc = CalculatePorc(.PorcInitial, .AddStep, .MaxStep)
                        .Duration = CalculateDuration(.Porc, .Level, .AddStep, .MaxStep)
                        
                     Case eBonusType.eObj
                        .ObjIndex = Bonus_NewObj(Level)
                        
                        If .ObjIndex <> -1 Then
                            .Duration = AdjustObjDuration(Level)
                            
                            
                        End If
                End Select
            End If
        End With
    Next A

    Bonus_GenerateRandomConfigs = Configs
End Function

' # Determina la duración del Objeto
Public Function AdjustObjDuration(ByVal Level As Integer) As Integer
    Dim durationInSeconds As Integer
    
    ' Definir intervalos de duración en segundos para diferentes niveles
    ' Puedes personalizar estos valores según tus necesidades
    Select Case Level
        Case 1
            ' Nivel 1: Entre 30 minutos y 2 horas (en segundos)
            durationInSeconds = RandomNumber(30 * 60, 2 * 60 * 60) ' 30 minutos a 2 horas
        Case 2
            ' Nivel 2: Entre 2 horas y 4 horas (en segundos)
            durationInSeconds = RandomNumber(2 * 60 * 60, 4 * 60 * 60) ' 2 horas a 4 horas
        Case 3
            ' Nivel 3: Entre 4 horas y 6 horas (en segundos)
            durationInSeconds = RandomNumber(4 * 60 * 60, 6 * 60 * 60) ' 4 horas a 6 horas
        Case 4
            ' Nivel 4: Entre 6 horas y 12 horas (en segundos)
            durationInSeconds = RandomNumber(6 * 60 * 60, 12 * 60 * 60) ' 6 horas a 12 horas
        Case Else
            ' Niveles superiores
            ' Usar intervalos de días para niveles más altos
            Dim minDays As Integer
            Dim maxDays As Integer
            
            Select Case Level
                Case 5
                    minDays = 1
                    maxDays = 2
                Case 6
                    minDays = 2
                    maxDays = 3
                Case 7
                    minDays = 3
                    maxDays = 4
                Case 8
                    minDays = 4
                    maxDays = 5
                Case Else
                    ' Nivel 9 y superior: Entre 5 y 7 días (en segundos)
                    minDays = 5
                    maxDays = 7
            End Select
            
            durationInSeconds = RandomNumber(minDays * 24 * 60 * 60, maxDays * 24 * 60 * 60) ' De minDays a maxDays días
    End Select
    
    ' Ajustar la duración para que sea múltiplo de 30 o 60 minutos
    Dim remainder As Integer
    Dim interval As Integer
    interval = RandomNumber(30 * 60, 60 * 60) ' Intervalo de 30 a 60 minutos
    remainder = durationInSeconds Mod interval
    If remainder > 0 Then
        durationInSeconds = durationInSeconds + (interval - remainder)
    End If
    
    ' Devolver la duración ajustada en segundos
    AdjustObjDuration = durationInSeconds
End Function

'# Obtiene un objeto aleatoreo según el Level requerido.
Public Function Bonus_NewObj(ByVal Level As Integer) As Integer
    Dim CandidatosIndices() As Integer
    Dim CandidatosCount As Integer
    Dim i As Integer
    
    ' Paso 1: Construir una lista de índices de candidatos que cumplan con las restricciones de nivel
    For i = LBound(ObjsReward) To UBound(ObjsReward)
        If ObjsReward(i).Level <= Level Then
            ReDim Preserve CandidatosIndices(1 To CandidatosCount + 1)
            CandidatosCount = CandidatosCount + 1
            CandidatosIndices(CandidatosCount) = i
        End If
    Next i
    
    ' Paso 2: Elegir aleatoriamente un índice de candidato de la lista
    If CandidatosCount > 0 Then
        Dim IndiceAleatorio As Integer
        IndiceAleatorio = RandomNumber(1, CandidatosCount)
        Bonus_NewObj = ObjsReward(CandidatosIndices(IndiceAleatorio)).ObjIndex
    Else
        ' En caso de que no haya candidatos que cumplan con las restricciones de nivel
        Bonus_NewObj = -1 ' Puedes devolver un valor especial o -1 para indicar que no se pudo obtener ningún objeto
    End If
End Function


' # Obtiene la configuración de nivel en función del nivel proporcionado
Public Function GetLevelConfig(ByVal Level As Integer) As BonusConfig
    Dim config As BonusConfig
    
    config.MinPorcInitial = RoundNumberToStep(1 + ((Level - 1) * 5), 5)
    config.MaxPorcInitial = RoundNumberToStep(25 + ((Level - 1) * 10), 10)
    If config.MaxPorcInitial > 100 Then config.MaxPorcInitial = 100
    config.MaxAddStep = Choose(RandomNumber(1, 2), 5, 10)
    config.MaxMaxStep = Level \ 2
    If config.MaxMaxStep < 1 Then config.MaxMaxStep = 1
  
    GetLevelConfig = config
End Function


' Esta función redondea un número al múltiplo más cercano del valor de "step"
Function RoundNumberToStep(ByVal number As Integer, ByVal step As Integer) As Integer
    RoundNumberToStep = Round(number / step) * step
End Function
' # Función para calcular Porc basado en otros campos
Public Function CalculatePorc(ByVal PorcInitial As Byte, ByVal AddStep As Byte, ByVal MaxStep As Byte) As Byte
    Dim Porc As Integer
    Dim i As Long
    
    Porc = PorcInitial
    
    For i = 1 To MaxStep
        Porc = Porc + AddStep
    Next i

    ' Redondear a múltiplos de AddStep
    Porc = (Porc \ AddStep) * AddStep

    ' Asegurarse de que Porc no supera el 100%
    If Porc > 100 Then Porc = 100

    CalculatePorc = Porc
End Function


' # Función para calcular Duration basado en otros campos
Public Function CalculateDuration(ByVal Porc As Byte, ByVal Level As Byte, ByVal AddStep As Byte, ByVal MaxStep As Byte) As Long
    Dim baseDuration As Double
    Dim tempDuration As Double
    
    ' Convertir MinDuration y MaxDuration a minutos
    Dim minDurationInMinutes As Double
    Dim maxDurationInMinutes As Double
    minDurationInMinutes = MinDuration / 60
    maxDurationInMinutes = MaxDuration / 60
    
    ' Calcular una duración base en minutos que sea proporcional al nivel y al porcentaje
    baseDuration = minDurationInMinutes + ((maxDurationInMinutes - minDurationInMinutes) * (Porc / 100#)) * (Level / 10#)
    
    ' Ajustar la duración en función de otros parámetros como AddStep y MaxStep
    tempDuration = baseDuration * ((AddStep / 10#) + 1#) * ((MaxStep / 5#) + 1#)
    
    ' Redondear al múltiplo más cercano de 30 minutos o 1 hora, según corresponda
    Dim interval As Double
    If tempDuration <= 2 * 60 Then ' Menos de 2 horas, redondear a 30 minutos
        interval = 30
    Else
        interval = 60 ' 2 horas o más, redondear a 1 hora
    End If
    
    tempDuration = Round(tempDuration / interval) * interval
    
    ' Asegurarse de que la duración esté dentro de los límites deseados en minutos
    If tempDuration < minDurationInMinutes Then
        tempDuration = minDurationInMinutes
    ElseIf tempDuration > maxDurationInMinutes Then
        tempDuration = maxDurationInMinutes
    End If
    
    CalculateDuration = CLng(tempDuration) * 60 ' Convertir la duración a segundos
End Function





' # Print Text
Private Sub WriteToTxtFile(ByVal FilePath As String, ByVal Text As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile() ' Encuentra un número de archivo disponible
    
    Open FilePath For Append As fileNumber ' Abre el archivo para agregar contenido
    Print #fileNumber, Text
    Close fileNumber ' Cierra el archivo
End Sub
' # Ejemplificamos un poco para ver que tal anda...
Public Sub Bonus_TestingGeneration()
    Dim A As Integer
    Dim testBonus() As tBonus
    Dim FilePath As String

    ' Definir la ruta del archivo aquí (cámbialo según tus necesidades)
    FilePath = DatPath & "Bonus.txt"


    testBonus = Bonus_GenerateRandomConfigs(1, 100, 1, 100)
    
    For A = LBound(testBonus) To UBound(testBonus)
        If testBonus(A).Tipo > 0 Then
            WriteToTxtFile FilePath, "Efecto: " & A
            WriteToTxtFile FilePath, Bonus_Convert_String(testBonus(A))
            WriteToTxtFile FilePath, "-----------------------------"
        End If
    Next A
End Sub

' # String Convert Bonus
' # String Convert Bonus - Usando Array para reducir repetición
Private Function Bonus_Convert_String(ByRef Bonus As tBonus) As String

    Dim BonusTypes As Variant
    Dim BonusStrings As Variant
    Dim i As Long
    
    BonusTypes = Array(eBonusType.eGld, eBonusType.eExp, eBonusType.eDsp, eBonusType.ePoints, eBonusType.eObj)
    BonusStrings = Array("Oro", "Experiencia", "DSP", "Puntos de Partida", "Objeto Usable: Puntos de Partida")

    For i = LBound(BonusTypes) To UBound(BonusTypes)
        If Bonus.Tipo = BonusTypes(i) Then
            If Bonus.Tipo = eObj Then
                Bonus_Convert_String = "Nuevo Objeto: " & ObjData(Bonus.ObjIndex).Name & "% durante " & SecondsToHMS(Bonus.Duration) & "."
            Else
                Bonus_Convert_String = "Nuevo Efecto: Bonus de " & BonusStrings(i) & " " & Bonus.Porc & "% durante " & SecondsToHMS(Bonus.Duration) & "."
            End If
            Exit Function
        End If
    Next i

    Bonus_Convert_String = "Tipo de Bonus no reconocido."

End Function



