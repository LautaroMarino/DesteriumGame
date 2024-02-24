Attribute VB_Name = "mFunctions"
Option Explicit

Public Function Porcentaje_Per_Value(ByVal Value As Long, ByVal ValueMax As Long) As Single
        On Error GoTo ErrHandler
    Dim Porc As Single

    If Value <= 0 Or ValueMax <= 0 Then Exit Function
    
    Porc = Value / ValueMax ' Porcentaje en formato decimal, de 0 a 1
    Porcentaje_Per_Value = Round(Porc, 2)
    
    Exit Function
ErrHandler:
End Function
Public Function Porcentaje_Per_Frags(ByVal Frags As Long, ByVal FragsMax As Long) As Single
    On Error GoTo ErrHandler
    
    Dim Porc As Single

    If Frags <= 0 Or FragsMax <= 0 Then Exit Function

    ' Convierte los frags y el máximo de frags a un número de punto flotante para poder calcular logaritmos
    Dim FloatFrags As Double
    FloatFrags = CDbl(Frags)

    Dim FloatFragsMax As Double
    FloatFragsMax = CDbl(FragsMax)

    ' Calcula el logaritmo base 10 de los frags y del máximo de frags
    Dim LogFrags As Double
    LogFrags = Log(FloatFrags) / Log(10)

    Dim LogFragsMax As Double
    LogFragsMax = Log(FloatFragsMax) / Log(10)

    ' La bonificación es la suma de los frags y el logaritmo de los frags,
    ' dividida por la suma del máximo de frags y el logaritmo del máximo de frags
    Porc = (FloatFrags + LogFrags) / (FloatFragsMax + LogFragsMax)

    Porcentaje_Per_Frags = Round(Porc, 2)
    
    Exit Function
ErrHandler:
End Function




Public Function Porcentaje_Per_Level_Log(ByVal Level As Integer) As Single

    On Error GoTo ErrHandler
    
    Dim Porc As Single

    If Level <= 0 Then Exit Function
    
    ' Convierte el nivel y el nivel máximo a un número de punto flotante para poder calcular logaritmos
    Dim FloatLevel As Double
    FloatLevel = CDbl(Level)

    Dim FloatMaxLevel As Double
    FloatMaxLevel = CDbl(STAT_MAXELV)

    ' Calcula el logaritmo base 10 del nivel y del nivel máximo
    Dim LogLevel As Double
    LogLevel = Log(FloatLevel) / Log(10)

    Dim LogMaxLevel As Double
    LogMaxLevel = Log(FloatMaxLevel) / Log(10)

    ' La bonificación es la suma del nivel y el logaritmo del nivel,
    ' dividida por la suma del nivel máximo y el logaritmo del nivel máximo
    Porc = (FloatLevel + LogLevel) / (FloatMaxLevel + LogMaxLevel)

    ' Añadir un bonus significativo para los niveles 45, 46 y 47
    Select Case Level
        Case 45
          '  Porc = Porc + 0.1 ' Ajustar este valor para cambiar el bonus adicional
        Case 46
         '   Porc = Porc + 0.2 ' Ajustar este valor para cambiar el bonus adicional
        Case 47
         '   Porc = Porc + 0.3 ' Ajustar este valor para cambiar el bonus adicional
    End Select

   ' Ajustar el porcentaje de bonificación para los niveles menores o iguales a 25
    If Level <= 35 Then
        Porc = Porc - 0.04 * (35 - Level) ' Este valor puede ser ajustado para reducir el porcentaje de bonificación para estos niveles
    'ElseIf Level > 35 Then
        'Porc = Porc + 0.01 * (Level - 35) ' Añade un bono adicional del 0.08 por cada nivel superior a 35
    End If
    
    If Porc <= 0 Then Exit Function
    Porcentaje_Per_Level_Log = Round(Porc, 2)
    
    Exit Function
ErrHandler:
    
End Function

Public Function Porcentaje_Per_Level(ByVal Level As Long, ByVal LevelMax As Long) As Single
    On Error GoTo ErrHandler
    
    Dim Porc As Single

    If Level <= 0 Or LevelMax <= 0 Then Exit Function
    
    Dim ExponentialFactor As Single
    ExponentialFactor = 2 ' Ajustar este valor para controlar cuánto más rápidamente aumentan los porcentajes en los niveles más altos.
    
    Porc = ((Level / LevelMax) ^ ExponentialFactor)
    Porcentaje_Per_Level = Round(Porc, 2)
    Exit Function
ErrHandler:
End Function

' Transforma un valor decimal (de 1 a 3) en un porcentaje (de 100% a 300%)
Public Function Transformar_En_Porcentaje(ByRef ValorDecimal As Single) As Single
    On Error GoTo ErrHandler
    
    ' Escala el valor decimal al rango de 100 a 300
    Transformar_En_Porcentaje = (ValorDecimal - 1) * 100 + 100 ' ValorDecimal = 1 --> Porcentaje = 100, ValorDecimal = 3 --> Porcentaje = 300
    
    ' Redondea el resultado a 0 decimales
    Transformar_En_Porcentaje = Round(Transformar_En_Porcentaje, 0)
    
    ' Comprueba si el resultado excede el 300%
    If Transformar_En_Porcentaje > 300 Then
        Transformar_En_Porcentaje = 300
    End If
    Exit Function
ErrHandler:
    
End Function

' Función que calcula las kills máximas que un usuario puede realizar, teniendo en cuenta los rounds.
Public Function MaxKills(ByVal TeamCant As Integer, ByVal CuposMax As Integer, ByVal Rounds As Integer, ByVal RoundsFinal As Integer) As Integer
    On Error GoTo ErrHandler
    
    ' Calcula el número de equipos
    Dim numTeams As Integer
    numTeams = CuposMax / TeamCant
    
    ' Calcula el número total de jugadores en los otros equipos
    Dim otherPlayers As Integer
    otherPlayers = CuposMax - TeamCant
    
    ' Las kills máximas por partida son simplemente el número total de jugadores en los otros equipos
    Dim maxKillsPerGame As Integer
    maxKillsPerGame = otherPlayers

    ' Calcula el total de rounds
    Dim totalRounds As Integer
    totalRounds = Rounds + RoundsFinal

    ' Las kills máximas totales son las kills máximas por partida multiplicadas por el total de rounds
    MaxKills = maxKillsPerGame * totalRounds
    Exit Function
ErrHandler:
End Function

Function EsPar(numero As Integer) As Boolean
    If numero Mod 2 = 0 Then
        EsPar = True ' El número es par
    Else
        EsPar = False ' El número es impar
    End If
End Function

