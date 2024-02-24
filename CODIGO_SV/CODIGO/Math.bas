Attribute VB_Name = "Math"
Option Explicit

Private Const VB_MIN_BYTE As Byte = 0
Private Const VB_MIN_INT As Integer = -32768
Private Const VB_MIN_LONG As Long = -2147483648#

Private Const VB_MAX_BYTE As Byte = 255
Private Const VB_MAX_INT As Integer = 32767
Private Const VB_MAX_LONG As Long = 2147483647

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2957795130823 '180 / Pi


Private Const MAX_HISTORY As Long = 10
Private GeneratedNumbers As New Collection

Public Function mini(ByVal A As Integer, ByVal B As Integer) As Integer
    If A > B Then
        mini = B
    Else
        mini = A
    End If
End Function

Public Function maxi(ByVal A As Integer, ByVal B As Integer) As Integer
    If A > B Then
        maxi = A
    Else
        maxi = B
    End If
End Function


Public Function minl(ByVal A As Long, ByVal B As Long) As Long
    If A > B Then
        minl = B
    Else
        minl = A
    End If
End Function

Public Function maxl(ByVal A As Long, ByVal B As Long) As Long
    If A > B Then
        maxl = A
    Else
        maxl = B
    End If
End Function
Public Function mins(ByVal A As Single, ByVal B As Single) As Single
    If A > B Then
        mins = B
    Else
        mins = A
    End If
End Function

Public Function maxs(ByVal A As Single, ByVal B As Single) As Single
    If A > B Then
        maxs = A
    Else
        maxs = B
    End If
End Function

Public Function CByteSeguro(ByVal expresion As String) As Byte
    Dim Temp As Single
    
    Temp = val(expresion)
    
    '¿Esta dentro de los limites?
    If Temp < VB_MIN_BYTE Or Temp > VB_MAX_BYTE Then
        CByteSeguro = VB_MAX_BYTE
    Else
        CByteSeguro = CByte(Temp)
    End If
End Function

Public Function Angulo(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Single
    If X2 - X1 = 0 Then
        If Y2 - Y1 = 0 Then
            Angulo = 90
        Else
            Angulo = 270
        End If
    Else
        Angulo = Atn((Y2 - Y1) / (X2 - X1)) * RadianToDegree
        If (X2 - X1) < 0 Or (Y2 - Y1) < 0 Then Angulo = Angulo + 180
        If (X2 - X1) > 0 And (Y2 - Y1) < 0 Then Angulo = Angulo - 180
        If Angulo < 0 Then Angulo = Angulo + 360
    End If
End Function


Public Function CIntSeguro(ByVal expresion As String) As Integer
    Dim Temp As Single
    
    Temp = val(expresion)
    
    '¿Esta dentro de los limites?
    If Temp < VB_MIN_INT Or Temp > VB_MAX_INT Then
        CIntSeguro = VB_MAX_INT
    Else
        CIntSeguro = CInt(Temp)
    End If
End Function

Sub AddtoVar(ByRef Var As Variant, ByVal Addon As Variant, ByVal max As Variant)
    'Le suma un valor a una variable respetando el maximo valor
    If Var >= max Then
        Var = max
    Else
        Var = Var + Addon
        If Var > max Then
            Var = max
        End If
    End If
End Sub

Public Sub RestToVar(ByRef Var As Integer, ByVal cantidad As Integer, ByVal Min As Integer)
    'Le suma un valor a una variable respetando el maximo valor
    Var = Var - cantidad
    If Var < Min Then
        Var = Min
    End If
End Sub
Function max(ByVal A As Double, ByVal B As Double) As Double
        On Error GoTo max_Err
100     If A > B Then
102         max = A
        Else
104         max = B
        End If
        Exit Function
max_Err:
End Function
Function Min(ByVal A As Double, ByVal B As Double) As Double
        On Error GoTo min_Err
100     If A < B Then
102         Min = A
        Else
104         Min = B
        End If
        Exit Function
min_Err:
End Function
Public Function Porcentaje(ByVal Total As Double, ByVal Porc As Double) As Double
        On Error GoTo Porcentaje_Err
100     Porcentaje = (Total * Porc) / 100
        Exit Function
Porcentaje_Err:
End Function
Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
        On Error GoTo Distancia_Err
100     Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100&)
        Exit Function
Distancia_Err:
End Function
Function Distance(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
        On Error GoTo Distance_Err
100     Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
        Exit Function
Distance_Err:
End Function
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        On Error GoTo RandomNumber_Err
        
        'Randomize
100     RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
        Exit Function
RandomNumber_Err:
End Function
Public Function RandomNumberPower(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    
    On Error GoTo ErrHandler
    
    RandomNumberPower = RandomNumber(LowerBound, UpperBound)
    
    Exit Function
    
    ' Inicializar generador de números aleatorios usando el valor de tiempo actual del sistema
    ' Esta línea se puede comentar si quieres evitar el uso del tiempo del sistema
     Randomize Timer

    ' Generar número aleatorio entre 0 (inclusive) y 1 (exclusivo)
    Dim rndValue As Double
    rndValue = Rnd()
    
    ' Escalar el valor al rango deseado (LowerBound a UpperBound)
    Dim Result As Long
    Result = LowerBound + (UpperBound - LowerBound) * rndValue

    ' Redondear el resultado para que sea un entero
    RandomNumberPower = Int(Result)
    Exit Function
ErrHandler:
    
End Function
