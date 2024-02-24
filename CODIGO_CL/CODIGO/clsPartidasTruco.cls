VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPartidasTruco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' # Bazas para ganar la mano
Private Const MAX_BAZAS As Integer = 2

' # 40 Cartas en una baraja
Private mBaraja As New clsCartasBaraja

' # Jugadores participes de la Partida
Private mJugadores() As String

' # Cartas que el jugador le toca en su mano
Private mManoOrig() As Collection

' # Cartas que el jugador va tirando
Private mManoJugadores() As Collection

' # Contabilidad de las manos ganadas
Private mBazasGanadas() As Integer

' # Monedas DSP que se apuesta
Private mMonedasDSP As Integer

' # Ronda
Private mRondaActual As Integer

' # Turno de
Private mTurnoActual As Integer

' # Puntos para ganar la partida
Private mPointsWon As Integer

' # Puntos de los jugadores
Private mPointsJugadores() As Integer

Public Property Get Jugadores() As String()
    Jugadores = mJugadores
End Property

Public Property Get MonedasDSP() As Integer
    MonedasDSP = mMonedasDSP
End Property

Public Property Let MonedasDSP(ByVal Value As Integer)
    mMonedasDSP = Value
End Property

Public Property Get RondaActual() As Integer
    RondaActual = mRondaActual
End Property

' # Inicializa una nueva partida
Public Sub IniciarPartida(Jugadores() As String, _
                          ByVal PointsWon As Integer, _
                          ByVal MonedasDSP As Integer)

    mJugadores = Jugadores
    mMonedasDSP = MonedasDSP
    mPointsWon = PointsWon
    ReDim mPointsJugadores(1 To UBound(mJugadores))
    mTurnoActual = 1
End Sub

' # Reparte las cartas a los personajes
Public Sub RepartirCartas()
    Dim i As Integer
    Dim Carta As clsCarta
    
    ' Barajar las cartas antes de repartir
    mBaraja.BarajarCartas

    ' Inicializar la mano privada, la mano de juego y las bazas ganadas de cada jugador como una colección
    ReDim mManoOrig(1 To UBound(mJugadores))
    ReDim mManoJugadores(1 To UBound(mJugadores))
    ReDim mBazasGanadas(1 To UBound(mJugadores))
    
    For i = 1 To UBound(mJugadores)
        Set mManoOrig(i) = New Collection
        Set mManoJugadores(i) = New Collection
        mBazasGanadas(i) = 0
    Next i

    ' Repartir tres cartas a cada jugador en orden
    For i = 1 To 3
        Dim j As Integer
        For j = 1 To UBound(mJugadores)

            ' Repartir cartas de juego
            Set Carta = mBaraja.ObtenerNuevaCarta
            Carta.Jugada = False
            mManoOrig(j).Add Carta

            Debug.Print mJugadores(j) & ": Privada - " & mManoOrig(j).Item(i).ObtenerCarta & ", En Juego - " & Carta.ObtenerCarta

        Next j
    Next i
End Sub

' # Selecciona una carta al Azar
' # Utilizado cuando supera el tiempo de respuesta o para testeo.
Public Function SeleccionarCartaAlAzar(ByVal Jugador As String) As clsCarta
    ' Obtener el índice del jugador
    Dim indiceJugador As Integer
    indiceJugador = ArrayIndex(mJugadores, Jugador)
    
    Randomize
    
    ' Elegir una carta al azar
    Dim cartaElegida As clsCarta
    Do
        Set cartaElegida = mManoOrig(indiceJugador)(Int((mManoOrig(indiceJugador).Count * Rnd) + 1))
    Loop While cartaElegida.Jugada
    

    ' Puedes hacer algo con la carta elegida, por ejemplo, mostrarla en la consola
    'Debug.Print Jugador & " seleccionó al azar: " & cartaElegida.ObtenerCarta
    
    Set SeleccionarCartaAlAzar = cartaElegida
End Function

' # El jugador realiza una Jugada
Public Sub RealizarJugada(ByVal Jugador As String, ByVal Carta As String)
    ' Obtener el índice del jugador
    Dim indiceJugador As Integer
    indiceJugador = ArrayIndex(mJugadores, Jugador)
    
    ' Almacenar la carta jugada en la mano de juego del jugador
    If Not mBaraja.ObtenerCartaPorNombre(Carta).Jugada Then
        mManoJugadores(indiceJugador).Add mBaraja.ObtenerCartaPorNombre(Carta)
        ' Marcar la carta como jugada
        mBaraja.ObtenerCartaPorNombre(Carta).MarcarComoJugada
    Else
        MsgBox "La carta ya fue jugada.", vbExclamation, "Error"
        ' Puedes agregar lógica adicional si es necesario
    End If
    
    mBaraja.ObtenerCartaPorNombre(Carta).Jugada = True

    ' Verificar si todos los jugadores han jugado en esta ronda
    If mTurnoActual = UBound(mJugadores) Then
        ' Avanzar a la siguiente ronda
        mRondaActual = mRondaActual + 1
        
        ' Calcular el ganador de la baza
        Dim ganadorBaza As String
        ganadorBaza = DeterminarGanadorBaza()
        
        Debug.Print "Baza " & mRondaActual & ": Ganador - " & ganadorBaza
                
        ' Sumar una baza al ganador
        mBazasGanadas(ArrayIndex(mJugadores, ganadorBaza)) = mBazasGanadas(ArrayIndex(mJugadores, ganadorBaza)) + 1
        
        ' Verificar si se han jugado todas las bazas de la mano para pasar a la siguiente ronda
        If mBazasGanadas(ArrayIndex(mJugadores, ganadorBaza)) >= ObtenerCantidadBazasParaMano Then
            ' Ganador de la MANO
            Debug.Print "MANO " & ": Ganador - " & ganadorBaza
            
            ' # Otorga los puntos
            mPointsJugadores(ArrayIndex(mJugadores, ganadorBaza)) = mPointsJugadores(ArrayIndex(mJugadores, ganadorBaza)) + 1
            
            mRondaActual = 0
            ' Repartir nuevas cartas
            Call RepartirCartas
        End If
        
        
        ' Comprobar si se alcanzaron los puntos necesarios para ganar
         If ObtenerPuntosJugador(ganadorBaza) >= mPointsWon Then
            MsgBox "¡Partida finalizada! El ganador es " & ganadorBaza, vbInformation, "Fin de la Partida"
            ' Puedes agregar más lógica aquí, como reiniciar la partida, mostrar estadísticas, etc.
            Exit Sub
        End If
    End If
    
    ' Avanzar al siguiente turno
    mTurnoActual = (mTurnoActual Mod UBound(mJugadores)) + 1
End Sub

Private Function ObtenerCantidadBazasParaMano() As Integer
    ' Devuelve la cantidad de bazas necesarias para completar una mano

    If mRondaActual = 2 Then
        ' En la segunda ronda, la PARDA en primera influye en la cantidad de bazas
        If SeJugoPardaEn(1) Then
            ObtenerCantidadBazasParaMano = 1
        Else
            ObtenerCantidadBazasParaMano = 2
        End If
    Else
        ' En la tercera ronda, se juegan 3 bazas
        ObtenerCantidadBazasParaMano = 2
    End If
End Function
' # Determina quien gana la mano segun los valores de sus cartas
Private Function DeterminarGanadorBaza() As String
    ' Lógica para determinar el ganador de la baza
    Dim ganador As String, i As Long
    Dim valorGanador As Integer

    ' Inicializar el valor ganador
    valorGanador = -1

    ' Iterar a través de las manos de juego de los jugadores para encontrar la carta ganadora
    For i = 1 To UBound(mJugadores)
        Dim ultimaCarta As clsCarta
        Set ultimaCarta = mManoJugadores(i).Item(mRondaActual)
    
        Debug.Print "Carta de " & mJugadores(i) & ": " & ultimaCarta.ObtenerCarta & " - Valor Numérico: " & ultimaCarta.ObtenerValorNumerico
    
        If ultimaCarta.ObtenerValorNumerico > valorGanador Then
            ' La carta actual es la ganadora
            valorGanador = ultimaCarta.ObtenerValorNumerico
            ganador = mJugadores(i)
        End If
    Next i


    DeterminarGanadorBaza = ganador
End Function

Private Function SeJugoPardaEn(ByVal RondaActual As Integer) As Boolean
    ' Devuelve True si se jugó PARDA en la primera ronda
    ' Utiliza la lógica de parda basada en la comparación de cartas
    
    ' Asumiendo que todos los jugadores tienen al menos una carta en la primera ronda
    Dim carta1 As clsCarta
    Dim carta2 As clsCarta
    
    ' Obtén las primeras cartas jugadas por dos jugadores (podrías ajustar esto según tu lógica)
    Set carta1 = mManoJugadores(1).Item(RondaActual)
    Set carta2 = mManoJugadores(2).Item(RondaActual)
    
    ' Devuelve True si las dos cartas son iguales (parda), False en caso contrario
    SeJugoPardaEn = mBaraja.EsParda(carta1, carta2)
End Function

Public Function ObtenerPuntosJugador(ByVal Jugador As String) As Integer
    ' Obtener los puntos del jugador especificado
    ' (puedes ajustar esta función según tus reglas)
    ObtenerPuntosJugador = mPointsJugadores(ArrayIndex(mJugadores, Jugador))
End Function
Private Function ArrayIndex(arr() As String, Value As String) As Integer
    ' Función auxiliar para obtener el índice de un elemento en un array
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = Value Then
            ArrayIndex = i
            Exit Function
        End If
    Next i
    ArrayIndex = -1 ' Valor por defecto si no se encuentra el elemento
End Function




