Attribute VB_Name = "Matematicas"
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

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Currency
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Porcentaje = (Total * Porc) / 100
End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)
End Function

Function Distance(ByVal X1 As Integer, _
                  ByVal Y1 As Integer, _
                  ByVal X2 As Integer, _
                  ByVal Y2 As Integer) As Double
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Encuentra la distancia entre dos puntos

    Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        '<EhHeader>
        On Error GoTo RandomNumber_Err
        '</EhHeader>
100     RandomNumber = Int(((UpperBound + 1) - LowerBound) * Rnd) + LowerBound
        '<EhFooter>
        Exit Function

RandomNumber_Err:
        LogError Err.description & vbCrLf & _
               "in RandomNumber " & _
               "at line " & Erl

        '</EhFooter>
End Function
