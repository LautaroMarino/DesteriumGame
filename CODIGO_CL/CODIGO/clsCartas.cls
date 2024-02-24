VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase Carta
Option Explicit

Private mPalo As String
Private mValor As String
Private mValorNumerico As Integer
Private mJugada As Boolean

Public Property Get palo() As String
    palo = mPalo
End Property

Public Property Let palo(ByVal Value As String)
    mPalo = Value
End Property

Public Property Get Valor() As String
    Valor = mValor
End Property

Public Property Let Valor(ByVal Value As String)
    mValor = Value
    ' Asignamos un valor numérico basado en el orden de las cartas en el truco
    mValorNumerico = ValorNumericoFromTexto(Value, mPalo)
End Property
Public Property Let Jugada(ByVal Value As Boolean)
    mJugada = Value
End Property

Public Property Get Jugada() As Boolean
    Jugada = mJugada
End Property

Public Sub MarcarComoJugada()
    mJugada = True
End Sub
Private Function ValorNumericoFromTexto(ByVal valorTexto As String, ByVal palo As String) As Integer
    ' Asignamos un valor numérico a cada carta basándonos en la jerarquía del truco
    Select Case valorTexto
        Case "1"
            If palo = "Espadas" Then
                ValorNumericoFromTexto = 14 ' As de Espadas
            ElseIf palo = "Bastos" Then
                ValorNumericoFromTexto = 13 ' As de Bastos
            ElseIf palo = "Oros" Or palo = "Copas" Then
                ValorNumericoFromTexto = 8 ' Anchos falsos
            End If
        Case "7"
            If palo = "Espadas" Then
                ValorNumericoFromTexto = 12 ' 7 de Espadas
            ElseIf palo = "Oros" Then
                ValorNumericoFromTexto = 11 ' 7 de Oros
            ElseIf palo = "Copas" Or palo = "Bastos" Then
                ValorNumericoFromTexto = 4 ' 7 de Copas y Bastos
            End If
        Case "3"
            ValorNumericoFromTexto = 10 ' Los 3 (cualquier palo)
        Case "2"
            ValorNumericoFromTexto = 9 ' Los 2 (cualquier palo)
        Case "12"
            ValorNumericoFromTexto = 7 ' Los 12 (reyes)
        Case "11"
            ValorNumericoFromTexto = 6 ' Los 11 (caballos)
        Case "10"
            ValorNumericoFromTexto = 5 ' Los 10 (sotas)
        Case "6"
            ValorNumericoFromTexto = 3 ' Los 6 (cualquier palo)
        Case "5"
            ValorNumericoFromTexto = 2 ' Los 5 (cualquier palo)
        Case "4"
            ValorNumericoFromTexto = 1 ' Los 4 (cualquier palo)
        Case Else
            ValorNumericoFromTexto = 0 ' Valor por defecto para casos no esperados
    End Select
End Function

Public Function ObtenerCarta() As String
    ObtenerCarta = mValor & " de " & mPalo
End Function

Public Function ObtenerValorNumerico() As Integer
    ObtenerValorNumerico = mValorNumerico
End Function
