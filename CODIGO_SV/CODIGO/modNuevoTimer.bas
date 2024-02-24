Attribute VB_Name = "modNuevoTimer"
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

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, _
                                            Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual
        End If

        IntervaloPermiteLanzarSpell = True
    Else
        IntervaloPermiteLanzarSpell = False
    End If

End Function
Public Function IntervaloPermiteShiftear(ByVal UserIndex As Integer, _
                                            Optional ByVal Actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerShiftear >= IntervaloUserPuedeShiftear Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerShiftear = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerShiftear = TActual
        End If

        IntervaloPermiteShiftear = True
    Else
        IntervaloPermiteShiftear = False
    End If

End Function

Public Function IntervaloPermiteCaspear(ByVal UserIndex As Integer, _
                                        Optional ByVal Actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.CaspeoTime >= 2000 Then
        If Actualizar Then
            UserList(UserIndex).Counters.CaspeoTime = TActual
        End If

        IntervaloPermiteCaspear = True
    Else
        IntervaloPermiteCaspear = False
    End If

End Function

Public Function IntervaloPermiteMoverUsuario(ByVal UserIndex As Integer, _
                                            Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteLanzarSpell_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime

102     If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimerLanzarSpell = TActual
                ' Actualizo spell-attack
108             UserList(UserIndex).Counters.TimerMagiaGolpe = TActual
            End If

110         IntervaloPermiteMoverUsuario = True
        Else
112         IntervaloPermiteMoverUsuario = False
        End If

        '<EhFooter>
        Exit Function

IntervaloPermiteLanzarSpell_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteLanzarSpell " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, _
                                       Optional ByVal Actualizar As Boolean = True) As Boolean
     '***************************************************
     'Author: Unknown
     'Last Modification: -
     '
     '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteAtacar_Err
        '</EhHeader>

     Dim TActual As Long

100  TActual = GetTime

102  If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
104      If Actualizar Then
106          UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
             ' Actualizo attack-spell
108          UserList(UserIndex).Counters.TimerGolpeMagia = TActual
             ' Actualizo attack-use
110          UserList(UserIndex).Counters.TimerGolpeUsar = TActual
         End If

112      IntervaloPermiteAtacar = True
     Else
114      IntervaloPermiteAtacar = False
     End If

        '<EhFooter>
        Exit Function

IntervaloPermiteAtacar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteAtacar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: ZaMa
        'Checks if the time that passed from the last hit is enough for the user to use a potion.
        'Last Modification: 06/04/2009
        '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteGolpeUsar_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime

102     If TActual - UserList(UserIndex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimerGolpeUsar = TActual
            End If

108         IntervaloPermiteGolpeUsar = True
        Else
110         IntervaloPermiteGolpeUsar = False
        End If

        '<EhFooter>
        Exit Function

IntervaloPermiteGolpeUsar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteGolpeUsar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloPermiteMagiaGolpe_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        Dim TActual As Long
    
100     With UserList(UserIndex)
        
102         TActual = GetTime
        
104         If TActual - .Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
106             If Actualizar Then
108                 .Counters.TimerMagiaGolpe = TActual
                End If

110             IntervaloPermiteMagiaGolpe = True
            Else
112             IntervaloPermiteMagiaGolpe = False
            End If

        End With

        '<EhFooter>
        Exit Function

IntervaloPermiteMagiaGolpe_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteMagiaGolpe " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteGolpeMagia_Err
        '</EhHeader>

        Dim TActual As Long
    
100     TActual = GetTime
    
102     If TActual - UserList(UserIndex).Counters.TimerGolpeMagia >= IntervaloGolpeMagia Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            End If

108         IntervaloPermiteGolpeMagia = True
        Else
110         IntervaloPermiteGolpeMagia = False
        End If

        '<EhFooter>
        Exit Function

IntervaloPermiteGolpeMagia_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteGolpeMagia " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, _
                                         Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteTrabajar_Err
        '</EhHeader>

        Dim TActual As Long
    
100     TActual = GetTime
    
102     If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
104         If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
106         IntervaloPermiteTrabajar = True
        Else
108         IntervaloPermiteTrabajar = False
        End If

        '<EhFooter>
        Exit Function

IntervaloPermiteTrabajar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteTrabajar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, _
                                     Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 25/01/2010 (ZaMa)
        '25/01/2010: ZaMa - General adjustments.
        '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteUsar_Err
        '</EhHeader>

        Dim TActual As Long
    
100     TActual = GetTime
    
102     If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimerUsar = TActual
108             UserList(UserIndex).Counters.failedUsageAttempts = 0
            End If

110         IntervaloPermiteUsar = True
        Else
112         IntervaloPermiteUsar = False
        
114         UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
            'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
116         If UserList(UserIndex).Counters.failedUsageAttempts = 10 Then
                'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Ip & " estuvo alterando el intervalo 'IntervaloPermiteUsar'", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
118             UserList(UserIndex).Counters.failedUsageAttempts = 0
            End If
        End If

        '<EhFooter>
        Exit Function

IntervaloPermiteUsar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteUsar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo IntervaloPermiteUsarArcos_Err
        '</EhHeader>

        Dim TActual As Long
    
100     TActual = GetTime
    
102     If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
104         If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
106         IntervaloPermiteUsarArcos = True
        Else
108         IntervaloPermiteUsarArcos = False
        End If

        '<EhFooter>
        Exit Function

IntervaloPermiteUsarArcos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteUsarArcos " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteSerAtacado(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = False) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloPermiteSerAtacado_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 13/11/2009
        '13/11/2009: ZaMa - Add the Timer which determines wether the user can be atacked by a NPc or not
        '**************************************************************
        Dim TActual As Long
    
100     TActual = GetTime
    
102     With UserList(UserIndex)

            ' Inicializa el timer
104         If Actualizar Then
106             .Counters.TimerPuedeSerAtacado = TActual
108             .flags.NoPuedeSerAtacado = True
110             IntervaloPermiteSerAtacado = False
            Else

112             If TActual - .Counters.TimerPuedeSerAtacado >= IntervaloPuedeSerAtacado Then
114                 .flags.NoPuedeSerAtacado = False
116                 IntervaloPermiteSerAtacado = True
                Else
118                 IntervaloPermiteSerAtacado = False
                End If
            End If

        End With

        '<EhFooter>
        Exit Function

IntervaloPermiteSerAtacado_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteSerAtacado " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPerdioNpc(ByVal UserIndex As Integer, _
                                   Optional ByVal Actualizar As Boolean = False) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloPerdioNpc_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 13/11/2009
        '13/11/2009: ZaMa - Add the Timer which determines wether the user still owns a Npc or not
        '**************************************************************
        Dim TActual As Long
    
100     TActual = GetTime
    
102     With UserList(UserIndex)

            ' Inicializa el timer
104         If Actualizar Then
106             .Counters.TimerPerteneceNpc = TActual
108             IntervaloPerdioNpc = False
            Else

110             If TActual - .Counters.TimerPerteneceNpc >= IntervaloOwnedNpc Then
112                 IntervaloPerdioNpc = True
                Else
114                 IntervaloPerdioNpc = False
                End If
            End If

        End With

        '<EhFooter>
        Exit Function

IntervaloPerdioNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPerdioNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Function



Public Function IntervaloGoHome(ByVal UserIndex As Integer, _
                                Optional ByVal TimeInterval As Long, _
                                Optional ByVal Actualizar As Boolean = False) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloGoHome_Err
        '</EhHeader>

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        '01/06/2010: ZaMa - Add the Timer which determines wether the user can be teleported to its home or not
        '**************************************************************
        Dim TActual As Long
    
100     TActual = GetTime
    
102     With UserList(UserIndex)

            ' Inicializa el timer
104         If Actualizar Then
106             .flags.Traveling = 1
108             .Counters.goHome = TActual + TimeInterval
            Else

110             If TActual >= .Counters.goHome Then
112                 IntervaloGoHome = True
114                 Call WriteUpdateGlobalCounter(UserIndex, 4, 0)
                End If
            End If

        End With

        '<EhFooter>
        Exit Function

IntervaloGoHome_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloGoHome " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteUsarClick(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloPermiteUsarClick_Err
        '</EhHeader>

        Dim TActual As Long

100     With UserList(UserIndex).Counters
102         TActual = GetTime()

104         If (TActual - UserList(UserIndex).Counters.TimerUsarClick) >= IntervaloUserPuedeUsarClick Then
106             If Actualizar Then
                    '.TimerUsar = TActual
108                 .TimerUsarClick = TActual

                End If

110             IntervaloPermiteUsarClick = True
            Else
112             IntervaloPermiteUsarClick = False
            
114             UserList(UserIndex).Counters.failedUsageAttempts_Clic = UserList(UserIndex).Counters.failedUsageAttempts_Clic + 1
        
                'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
116             If UserList(UserIndex).Counters.failedUsageAttempts_Clic = 10 Then
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Account.Sec.IP_Address & " estuvo alterando el intervalo 'IntervaloPermiteUsar'", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
118                 UserList(UserIndex).Counters.failedUsageAttempts_Clic = 0
                End If
            
            End If

        End With

        '<EhFooter>
        Exit Function

IntervaloPermiteUsarClick_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteUsarClick " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Interval_Drop(ByVal UserIndex As Integer, _
                              Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Interval_Drop_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime()

102     If TActual - UserList(UserIndex).Counters.TimeDrop >= IntervalDrop Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimeDrop = TActual
            End If

108         Interval_Drop = True
        Else
110         Interval_Drop = False
        End If

        '<EhFooter>
        Exit Function

Interval_Drop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_Drop " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Interval_InfoChar(ByVal UserIndex As Integer, _
                                  Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Interval_InfoChar_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime()

102     If TActual - UserList(UserIndex).Counters.TimeInfoChar >= 10000 Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimeInfoChar = TActual
            End If

108         Interval_InfoChar = True
        Else
110         Interval_InfoChar = False
        End If

        '<EhFooter>
        Exit Function

Interval_InfoChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_InfoChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Interval_Commerce(ByVal UserIndex As Integer, _
                                  Optional ByVal Actualizar As Boolean = True) As Boolean

    
        
          '<EhHeader>
    On Error GoTo Interval_Commerce_Err
        '</EhHeader>
        
    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.TimeCommerce >= IntervalCommerce Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeCommerce = TActual
        End If

        Interval_Commerce = True
    Else
        Interval_Commerce = False
    End If
   '<EhFooter>
        Exit Function

Interval_Commerce_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_Commerce " & _
               "at line " & Erl
        
End Function

Public Function Interval_Message(ByVal UserIndex As Integer, _
                                 Optional ByVal Actualizar As Boolean = True) As Boolean

    
           '<EhHeader>
        On Error GoTo Interval_Message_Err
        '</EhHeader>
        
    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.TimeMessage >= IntervalMessage Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeMessage = TActual
        End If

        Interval_Message = True
    Else
        Interval_Message = False
    End If


        '<EhFooter>
        Exit Function

Interval_Message_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_Message " & _
               "at line " & Erl
        
End Function

Public Function Interval_Packet250(ByVal UserIndex As Integer, _
                                 Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Interval_Packet250_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime()

102     If TActual - UserList(UserIndex).Counters.Packet250 >= 250 Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.Packet250 = TActual
            End If

108         Interval_Packet250 = True
        Else
110         Interval_Packet250 = False
        End If

        '<EhFooter>
        Exit Function

Interval_Packet250_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_Packet250 " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Interval_Packet500(ByVal UserIndex As Integer, _
                                 Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Interval_Packet500_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime()

102     If TActual - UserList(UserIndex).Counters.Packet500 >= 500 Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.Packet500 = TActual
            End If

108         Interval_Packet500 = True
        Else
110         Interval_Packet500 = False
        End If

        '<EhFooter>
        Exit Function

Interval_Packet500_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_Packet500 " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Interval_Mao(ByVal UserIndex As Integer, _
                             Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Interval_Mao_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime

102     If TActual - UserList(UserIndex).Counters.TimeInfoMao >= IntervalInfoMao Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimeInfoMao = TActual
            End If

108         Interval_Mao = True
        Else
110         Interval_Mao = False
        End If

        '<EhFooter>
        Exit Function

Interval_Mao_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Interval_Mao " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function Interval_Equipped(ByVal UserIndex As Integer, _
                             Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Interval_Equipped_Err
        '</EhHeader>

        Dim TActual As Long

100     TActual = GetTime

102     If TActual - UserList(UserIndex).Counters.TimeEquipped >= IntervaloEquipped Then
104         If Actualizar Then
106             UserList(UserIndex).Counters.TimeEquipped = TActual
            End If

108         Interval_Equipped = True
        Else
110         Interval_Equipped = False
        End If

        '<EhFooter>
        Exit Function

Interval_Equipped_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Interval_Equipped.Interval_Mao " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function checkInterval(ByRef startTime As Long, _
                              ByVal timeNow As Long, _
                              ByVal interval As Long) As Boolean
        '<EhHeader>
        On Error GoTo checkInterval_Err
        '</EhHeader>

        Dim lInterval As Long

100     If timeNow < startTime Then
102         lInterval = &H7FFFFFFF - startTime + timeNow + 1
        Else
104         lInterval = timeNow - startTime
        End If

106     If lInterval >= interval Then
108         startTime = timeNow
110         checkInterval = True
        Else
112         checkInterval = False
        End If

        '<EhFooter>
        Exit Function

checkInterval_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.checkInterval " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPuedeRecibirAtaqueCriature(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloPuedeRecibirAtaqueCriature_Err
        '</EhHeader>

    Dim TActual                     As Long

100     If haciendoBK Then Exit Function

102     TActual = GetTime

104     With UserList(UserIndex).Counters
106         If TActual - .TimerPuedeRecibirAtaqueCriature >= 800 Then
108             If Actualizar Then
110                 .TimerPuedeRecibirAtaqueCriature = TActual
                End If

112             IntervaloPuedeRecibirAtaqueCriature = True
            Else
114             IntervaloPuedeRecibirAtaqueCriature = False
            End If
        End With

        '<EhFooter>
        Exit Function

IntervaloPuedeRecibirAtaqueCriature_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPuedeRecibirAtaqueCriature " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function IntervaloPermiteCastear(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo IntervaloPermiteCastear_Err
        '</EhHeader>

    Dim TActual                     As Long

100     If haciendoBK Then Exit Function

102     TActual = GetTime

104     With UserList(UserIndex).Counters
106         If TActual - .TimerPuedeCastear >= IntervaloPuedeCastear Then
108             If Actualizar Then
110                 .TimerPuedeCastear = TActual
                End If

112             IntervaloPermiteCastear = True
            Else
114             IntervaloPermiteCastear = False
            
116             .failedUsageCastSpell = .failedUsageCastSpell + 1
        
                'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
118             If .failedUsageCastSpell = 10 Then
                    'Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Ip & " estuvo alterando el intervalo 'IntervaloPuedeCastear'")
                    'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Ip & " estuvo alterando el intervalo 'IntervaloPuedeCastear'", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
120                 .failedUsageCastSpell = 0
                End If
            End If
        End With

        '<EhFooter>
        Exit Function

IntervaloPermiteCastear_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.IntervaloPermiteCastear " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function Intervalo_BotUseItem(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Intervalo_BotUseItem_Err
        '</EhHeader>
    
        Dim TActual As Long
100     TActual = GetTime
    
102     With Npclist(NpcIndex)
    
104         If TActual - .Contadores.UseItem >= BotIntelligence_Balance_UseItem(.Stats.Elv) Then
106             If Update Then
108                 .Contadores.UseItem = TActual
                End If
            
110             Intervalo_BotUseItem = True
            Else
112             Intervalo_BotUseItem = False
            End If
        
        End With
        '<EhFooter>
        Exit Function

Intervalo_BotUseItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Intervalo_BotUseItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function Intervalo_CriatureVelocity(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Intervalo_CriatureVelocity_Err
        '</EhHeader>
    
        Dim TActual As Long
100     TActual = GetTime
    
102     With Npclist(NpcIndex)
    
104         If TActual - .Contadores.Velocity >= .Velocity Then
106             If Update Then
108                 .Contadores.Velocity = TActual
                End If
            
110             Intervalo_CriatureVelocity = True
            Else
112             Intervalo_CriatureVelocity = False
            End If
        
        End With
        '<EhFooter>
        Exit Function

Intervalo_CriatureVelocity_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureVelocity " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Public Function Intervalo_CriatureAttack(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Intervalo_CriatureAttack_Err
        '</EhHeader>
    
        Dim TActual As Long
100     TActual = GetTime
    
102     With Npclist(NpcIndex)
              
104         If TActual - .Contadores.Attack >= .IntervalAttack Then
106             If Update Then
108                 .Contadores.Attack = TActual
                End If
            
110             Intervalo_CriatureAttack = True
            End If
        
        End With
        '<EhFooter>
        Exit Function

Intervalo_CriatureAttack_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureAttack " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Intervalo_CriatureDescanso(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Intervalo_CriatureDescanso_Err
        '</EhHeader>
    
        Dim TActual As Long
100     TActual = GetTime
    
102     With Npclist(NpcIndex)
    
104         If TActual - .Contadores.Descanso >= 30000 Then
106             If Update Then
108                 .Contadores.Descanso = TActual
                End If
            
110             Intervalo_CriatureDescanso = True
            End If
        
        End With
        '<EhFooter>
        Exit Function

Intervalo_CriatureDescanso_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureDescanso " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function Intervalo_CriatureMovimientoConstante(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean
        '<EhHeader>
        On Error GoTo Intervalo_CriatureMovimientoConstante_Err
        '</EhHeader>
    
        Dim TActual As Long
100     TActual = GetTime
    
102     With Npclist(NpcIndex)
    
104         If TActual - .Contadores.MovimientoConstante >= 10000 Then
106             If Update Then
108                 .Contadores.MovimientoConstante = TActual
                End If
            
110             Intervalo_CriatureMovimientoConstante = True
            End If
        
        End With
        '<EhFooter>
        Exit Function

Intervalo_CriatureMovimientoConstante_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureMovimientoConstante " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

