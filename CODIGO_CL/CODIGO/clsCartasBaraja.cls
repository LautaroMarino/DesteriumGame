VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartasBaraja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase BarajaTruco
Option Explicit

Private mCartas() As clsCarta
Private mCartasRepartidas As Integer
Private mPointsJugadores As Collection ' Usamos una Collection para almacenar puntos por jugador

Private Sub Class_Initialize()

    Set mPointsJugadores = New Collection
 
    CargarCartasTruco
End Sub

Private Sub CargarCartasTruco()
    Dim palos() As String
    Dim valores() As String
    Dim i As Integer
    Dim j As Integer
    Dim index As Integer

    palos = Split("Oros,Copas,Espadas,Bastos", ",")
    valores = Split("1,2,3,4,5,6,7,10,11,12", ",")

    ReDim mCartas(1 To (UBound(palos) + 1) * (UBound(valores) + 1))

    Debug.Print "Array Size: " & UBound(mCartas)  ' Add this line for debugging

    index = 1
    For i = LBound(palos) To UBound(palos)
        For j = LBound(valores) To UBound(valores)
            Debug.Print "i: " & i & ", j: " & j & ", index: " & index  ' Add this line for debugging

            Set mCartas(index) = New clsCarta
            mCartas(index).palo = palos(i)
            mCartas(index).Valor = valores(j)
            index = index + 1
        Next j
    Next i
    
    mCartasRepartidas = 0
End Sub




Public Function ObtenerTodasLasCartas() As clsCarta()
    ObtenerTodasLasCartas = mCartas
End Function

Public Function CompararCartas(ByVal carta1 As clsCarta, ByVal carta2 As clsCarta) As Integer
    ' Devuelve 1 si carta1 es mayor, -1 si carta2 es mayor, y 0 si son iguales
    Dim valorCarta1 As Integer
    Dim valorCarta2 As Integer

    valorCarta1 = carta1.ObtenerValorNumerico
    valorCarta2 = carta2.ObtenerValorNumerico

    If valorCarta1 > valorCarta2 Then
        CompararCartas = 1
    ElseIf valorCarta1 < valorCarta2 Then
        CompararCartas = -1
    Else
        CompararCartas = 0
    End If
End Function

Public Sub BarajarCartas()
    ' Algoritmo de barajado Fisher-Yates
    Dim i As Integer
    Dim j As Integer
    Dim temp As clsCarta

    For i = UBound(mCartas) To 2 Step -1
        j = Int((i - 1) * Rnd + 1)
        Set temp = mCartas(i)
        Set mCartas(i) = mCartas(j)
        Set mCartas(j) = temp
    Next i
    
     mCartasRepartidas = 0
End Sub
Public Function ObtenerNuevaCarta() As clsCarta
    ' Devuelve la próxima carta disponible
    If mCartasRepartidas < UBound(mCartas) Then
        mCartasRepartidas = mCartasRepartidas + 1
        Set ObtenerNuevaCarta = mCartas(mCartasRepartidas)

    Else
        Set ObtenerNuevaCarta = Nothing ' No hay más cartas disponibles
    End If
End Function

Public Function ObtenerCartaPorNombre(ByVal Nombre As String) As clsCarta
    ' Obtener una carta por su nombre
    ' (puedes ajustar esta función según tus reglas)
    Dim i As Integer
    For i = 1 To UBound(mCartas)
        If mCartas(i).ObtenerCarta = Nombre Then
            Set ObtenerCartaPorNombre = mCartas(i)
            Exit Function
        End If
    Next i
    ' Si no se encuentra la carta, devuelve Nothing
    Set ObtenerCartaPorNombre = Nothing
End Function
Public Function EsParda(ByVal carta1 As clsCarta, ByVal carta2 As clsCarta) As Boolean
    ' Devuelve True si las cartas son iguales (PARDA), False en caso contrario
    EsParda = (CompararCartas(carta1, carta2) = 0)
End Function
