Attribute VB_Name = "cTimer"
Option Explicit

''
' How many timers we are going to use-
'
' @see See MainTimer.CreateAll
Private Const CANTIDADTIMERS As Byte = 10

''
' A Timer data structure.
'
' @param Interval How long, in miliseconds, a cicle lasts.
' @param CurrentTick Current Tick in which the Timer is.
' @param StartTick Tick in which current cicle has started.
' @param Run True if the timer is active.

Private Type Timer

    Interval As Long
    CurrentTick As Long
    StartTick As Long
    Run As Boolean

End Type

'Timers
Public Timer(1 To CANTIDADTIMERS) As Timer

''
' Timer큦 Index.
'
' @param Attack                 Controls the Combat system.
' @param Work                   Controls the Work system.
' @param UseItemWithU           Controls the usage of items with the "U" key.
' @param UseItemWithDblClick    Controls the usage of items with double click.
' @param SendRPU                Controls the use of the "L" to request a pos update.
' @param CastSpell              Controls the casting of spells.
' @param Arrows                 Controls the shooting of arrows.
Public Enum TimersIndex

    Attack = 1
    Work = 2
    UseItemWithU = 3
    UseItemWithDblClick = 4
    SendRPU = 5
    CastSpell = 6
    Arrows = 7
    CastAttack = 8
    Packet250 = 9
    Packet500 = 10

End Enum

'Very percise counter 64bit system counter
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Function GetSystemTime() As Long
    Static Frequency As Currency
    Static offset    As Currency
    
    ' Lazy initialization of timer frequency and offset
    If (Frequency = 0) Then
        Call QueryPerformanceFrequency(Frequency)
        Call QueryPerformanceCounter(offset)
        
        GetSystemTime = 0
    Else
        Dim Value As Currency
        Call QueryPerformanceCounter(Value)
        
        GetSystemTime = ((Value - offset) / Frequency * 1000)
    End If
End Function
Public Sub LoadTimerIntervals()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    'Set the intervals of timers
    '***************************************************
    Call SetInterval(TimersIndex.Attack, IntervaloUserPuedeAtacar)
    Call SetInterval(TimersIndex.CastSpell, IntervaloUserPuedeCastear)
    Call SetInterval(TimersIndex.CastAttack, IntervaloGolpeMagia)
    
    'Call SetInterval(TimersIndex.Work, IntervaloUserPuedeAtacar)
    'Call SetInterval(TimersIndex.UseItemWithU, IntervaloUserPuedeAtacar)
    'Call SetInterval(TimersIndex.UseItemWithDblClick, IntervaloUserPuedeAtacar)
    'Call SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    'Call SetInterval(TimersIndex.Arrows, INT_ARROWS)
    '
    'Call SetInterval(TimersIndex.Packet250, INT_PACKET250)
    'Call SetInterval(TimersIndex.Packet500, INT_PACKET500)
End Sub

''
' Window큦 API Function.
' A milisecond pricision counter.
'
' @return   Miliseconds since midnight.


''
' Sets a new interval for a timer.
'
' @param TimerIndex Timer큦 Index
' @param Interval New lenght for the Timer큦 cicle in miliseconds.
' @remarks  Must be done after creating the timer and before using it, otherwise, Interval will be 0

Public Sub SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Long)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified:
    'Desc: Sets a new interval for a timer.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Interval = Interval
End Sub

''
' Check if the timer has already completed it큦 cicle.
'
' @param TimerIndex Timer큦 Index
' @param Restart If true, restart if we timer has ticked
' @return True if the interval has already been elapsed
' @remarks  Can큧 be done if the timer is stoped or if it had never been started.

Public Function Check(ByVal TimerIndex As TimersIndex, _
                      Optional Restart As Boolean = True) As Boolean

    '*************************************************
    'Author: Nacho Agustin (Integer)
    'Last modified: 08/26/06
    'Modification: NIGO: Added Restart as boolean
    'Desc: Checks if the Timer has alredy "ticked"
    'Returns: True if it has ticked, False if not.
    '*************************************************
    Timer(TimerIndex).CurrentTick = GetSystemTime - Timer(TimerIndex).StartTick 'Calcutates CurrentTick
    
    If Timer(TimerIndex).CurrentTick >= Timer(TimerIndex).Interval Then
        Check = True 'We have Ticked!

        If Restart Then
            Timer(TimerIndex).StartTick = GetSystemTime 'Restart Timer (Nicer than calling Restart() )

            If (TimerIndex = TimersIndex.Attack) Or (TimerIndex = TimersIndex.CastSpell) Then
                Timer(TimersIndex.CastAttack).StartTick = GetSystemTime 'Set Cast-Attack interval
            ElseIf TimerIndex = TimersIndex.CastAttack Then
                Timer(TimersIndex.Attack).StartTick = GetSystemTime 'Restart Attack interval
                Timer(TimersIndex.CastSpell).StartTick = GetSystemTime 'Restart Magic interval
            End If
        End If
    End If

End Function

''
' Restarts timer.
'
' @param TimerIndex Timer큦 Index

Public Sub Restart(ByVal TimerIndex As TimersIndex)

    '*************************************************
    'Author: Nacho Agustin (Integer)
    'Last modified:
    'Desc: Restarts timer
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).StartTick = GetSystemTime
End Sub

