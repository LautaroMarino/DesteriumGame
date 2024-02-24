Attribute VB_Name = "Mod_Balance"
Option Explicit

Public Type tRango

    minimo As Integer
    maximo As Integer

End Type

Public Type ModRaza

    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single

End Type

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel

Public Const AdicionalHPCazador  As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef        As Byte = 15

Public Const AumentoStBandido    As Byte = AumentoSTDef + 3

Public Const AumentoSTLadron     As Byte = AumentoSTDef + 3

Public Const AumentoSTMago       As Byte = AumentoSTDef - 1

Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25

Public ModRaza(1 To NUMRAZAS)             As ModRaza


Public Sub Load_Balance()
    
    Dim I As Long
    
    'Modificadores de Raza
    For I = 1 To NUMRAZAS

        With ModRaza(I)
            .Fuerza = 18 + Val(GetVar(IniPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Fuerza"))
            .Agilidad = 18 + Val(GetVar(IniPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Agilidad"))
            .Inteligencia = 18 + Val(GetVar(IniPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Inteligencia"))
            .Carisma = 18 + Val(GetVar(IniPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Carisma"))
            .Constitucion = 18 + Val(GetVar(IniPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Constitucion"))
        End With

    Next I
End Sub


Public Function getVidaIdeal(ByVal Elv As Byte, ByVal Class As Byte, ByVal Constitucion As Byte) As Single
        '<EhHeader>
        On Error GoTo getVidaIdeal_Err
        '</EhHeader>

        Dim promedio     As Single

        Dim vidaBase     As Integer

        Dim rangoAumento As tRango
    
100     vidaBase = 20 '+ Int(getPromedioAumentoVida(Class, Constitucion) + 0.5)
    
102     rangoAumento = getRangoAumentoVida(Class, Constitucion)
104     promedio = (rangoAumento.minimo + rangoAumento.maximo) / 2
    
106     getVidaIdeal = vidaBase + (Elv - 1) * promedio

        '<EhFooter>
        Exit Function

getVidaIdeal_Err:
        LogError err.Description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.getVidaIdeal " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

' Retrona el minimo/maximo de puntos de vida que pude subir este usuario por nivel.
Public Function getRangoAumentoVida(ByVal Class As Byte, ByVal Constitucion As Byte) As tRango

    getRangoAumentoVida.maximo = 0
    getRangoAumentoVida.minimo = 0

    Select Case Class

        Case eClass.Warrior

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 9
                    getRangoAumentoVida.maximo = 12

                Case 20
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 12

                Case 19
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 11

                Case 18
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 11

                Case Else
                    getRangoAumentoVida.minimo = 6 + AdicionalHPCazador
                    getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

            End Select

        Case eClass.Hunter
    
            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 9
                    getRangoAumentoVida.maximo = 11

                Case 20
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 11

                Case 19
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 11

                Case 18
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 11

                Case Else
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

            End Select

        Case eClass.Paladin

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 9
                    getRangoAumentoVida.maximo = 11

                Case 20
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 11

                Case 19
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 11

                Case 18
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 11

                Case Else
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

            End Select

        Case eClass.Thief

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 20
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case 19
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 8

                Case 16, 17
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = 7

                Case 16
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = 6

                Case 14
                    getRangoAumentoVida.minimo = 2
                    getRangoAumentoVida.maximo = 6

                Case 13
                    getRangoAumentoVida.minimo = 2
                    getRangoAumentoVida.maximo = 5

                Case 12
                    getRangoAumentoVida.minimo = 1
                    getRangoAumentoVida.maximo = 5

                Case 11
                    getRangoAumentoVida.minimo = 1
                    getRangoAumentoVida.maximo = 4

                Case 10
                    getRangoAumentoVida.minimo = 0
                    getRangoAumentoVida.maximo = 4

                Case Else
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPGuerrero

            End Select
    
        Case eClass.Mage

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 8

                Case 20
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 8

                Case 19
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 8

                Case 18
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = 8

                Case Else
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPGuerrero

            End Select


        Case eClass.Cleric

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

        Case eClass.Druid

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select
        
        Case eClass.Assasin

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

        Case eClass.Bard

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

        Case Else

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 20
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case 19
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 8

                Case Else
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

    End Select

End Function

Public Function Balance_AumentoMANA(ByVal Class As Byte, ByVal Raze As Byte) As Integer

        ' Aumento de maná según clase
        '<EhHeader>
        On Error GoTo Balance_AumentoMANA_Err

        '</EhHeader>
    
        Dim UserInteligencia As Byte

        Dim Elv              As Byte

        Dim A                As Long
        
        Dim TempMan As Long
        
        Elv = 47
        
100     UserInteligencia = ModRaza(Raze).Inteligencia
    
        On Error GoTo Balance_AumentoMANA_Error
        
        For A = 2 To Elv

102         Select Case Class
                    
                Case eClass.Paladin
                      
104                 TempMan = TempMan + UserInteligencia
                         
106             Case eClass.Mage
                      
108                 If Raze = Enano Then
110                     TempMan = TempMan + 2 * UserInteligencia
112                 ElseIf (TempMan >= 2000) Then
114                     TempMan = TempMan + (3 * UserInteligencia) / 2
                    Else
116                     TempMan = TempMan + 3 * UserInteligencia

                    End If
                    
                    If A = 2 Then
                        TempMan = TempMan + 103
                    End If
                   
118             Case eClass.Druid, eClass.Bard, eClass.Cleric
120                 TempMan = TempMan + (2 * UserInteligencia)
                  
                  If A = 2 Then
                    TempMan = TempMan + 50
                  End If
                  
122             Case eClass.Assasin
124                 TempMan = TempMan + UserInteligencia
                  
                  If A = 2 Then
                        A = 20
                  End If
                  
126             Case Else
128                 TempMan = 0

            End Select
        
        Next A
        
        
        Balance_AumentoMANA = TempMan

        On Error GoTo Balance_AumentoMANA_Err

        Exit Function

Balance_AumentoMANA_Error:

130     LogError "Error " & err.Number & " (" & err.Description & ") in procedure Balance_AumentoMANA of Módulo mBalance in line " & Erl

        '<EhFooter>
        Exit Function

Balance_AumentoMANA_Err:
        LogError err.Description & vbCrLf & "in ServidorArgentum.Mod_Balance.Balance_AumentoMANA " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Function

