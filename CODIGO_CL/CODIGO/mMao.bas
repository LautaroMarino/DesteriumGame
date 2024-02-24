Attribute VB_Name = "mMao"
Option Explicit

Public Sub Mercader_OrdenClass()

    Dim A    As Long, b As Long
    Dim Temp As tMercader
    
    For A = 1 To MERCADER_MAX_LIST
        For b = 1 To MERCADER_MAX_LIST - A

            With MercaderList(b)
                If .Chars(1).Class < MercaderList(b + 1).Chars(1).Class Then
                    Temp = MercaderList(b)
                    MercaderList(b) = MercaderList(b + 1)
                    MercaderList(b + 1) = Temp
                End If
            End With
            
            DoEvents
        Next b
        
        DoEvents
    Next A
                
End Sub
Public Sub Mercader_OrdenLevel()

    Dim A    As Long, b As Long
    Dim Temp As tMercader
    
    For A = 1 To MERCADER_MAX_LIST
        For b = 1 To MERCADER_MAX_LIST - A

            With MercaderList(b)
                If .Chars(1).Elv < MercaderList(b + 1).Chars(1).Elv Then
                    Temp = MercaderList(b)
                    MercaderList(b) = MercaderList(b + 1)
                    MercaderList(b + 1) = Temp
                End If
            End With
            
            DoEvents
        Next b
        
        DoEvents
    Next A
                
End Sub
Public Function Mercader_GenerateText(ByRef Char As tMercaderChar, _
                                      Optional ByVal ShortText As Boolean = False) As String

    On Error GoTo ErrHandler
    
    Dim Temp   As String
    Dim TempUP As Single
    
    
    With Char
        TempUP = UserCheckPromedy(.Elv, .Hp, .Class, .Constitucion)
        
        .DescShort = ListaClases(.Class) & " " & ListaRazasShort(.Raze) & " " & CStr(.Elv) & IIf(TempUP > 0, " +", vbNullString) & CStr(TempUP)
        
        If .Elv <> STAT_MAXELV Then
            .DescShort = .DescShort & " (" & Round(CDbl(.Exp) * CDbl(100) / CDbl(.Elu), 0) & "%)"
        End If
            
        .Desc = "» " & .DescShort
    
    End With
    
  
    Exit Function

ErrHandler:

End Function

Public Function Mercader_GenerateText1(ByVal MercaderSlot As Integer, _
                                      ByVal Slot As Byte, _
                                      Optional ByVal ShortText As Boolean = False) As String

    On Error GoTo ErrHandler
    
    Dim Temp   As String
    Dim TempUP As Single
    
    With MercaderList_Copy(MercaderSlot).Chars(Slot)
        TempUP = UserCheckPromedy(.Elv, .Hp, .Class, .Constitucion)
        
        .DescShort = ListaClases(.Class) & " " & ListaRazasShort(.Raze) & " " & CStr(.Elv) & IIf(TempUP > 0, " +", vbNullString) & CStr(TempUP)
        
        If .Elv <> STAT_MAXELV Then
            .DescShort = .DescShort & " (" & Round(CDbl(.Exp) * CDbl(100) / CDbl(.Elu), 0) & "%)"
        End If
            
        'UCase$(.Name) &
        .Desc = "» " & .DescShort
        
    End With
    
  
    Exit Function

ErrHandler:

End Function

Public Function UserCheckPromedy(ByVal Elv As Byte, ByVal Hp As Integer, ByVal Class As eClass, ByVal UserConstitucion As Byte) As Single
    
    Dim LvlReal    As Long

    Dim HpIdeal    As Single

    Dim Diferencia As Single
        
    LvlReal = (Elv - 1)
    HpIdeal = ((Balance_AumentoHP_Initial(Class, UserConstitucion) * LvlReal) + 20)
        
    UserCheckPromedy = (Hp - HpIdeal)
    
End Function


Public Function Balance_AumentoHP_Initial(ByVal Class As eClass, ByVal UserConstitucion As Byte) As Single

    On Error GoTo Balance_AumentoHP_Initial_Error
    
    Select Case Class

        Case eClass.Warrior

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 10.5 'RandomNumber(9, 12)

                Case 20: Balance_AumentoHP_Initial = 10 'RandomNumber(8, 12)

                Case 19: Balance_AumentoHP_Initial = 9.5 'RandomNumber(8, 11)

                Case 18: Balance_AumentoHP_Initial = 9 ' RandomNumber(7, 11)

                Case Else: Balance_AumentoHP_Initial = 8 'RandomNumber(6, UserConstitucion \ 2) + AdicionalHPGuerrero
            End Select

        Case eClass.Hunter

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 10 'RandomNumber(9, 11)

                Case 20: Balance_AumentoHP_Initial = 9.5 'RandomNumber(8, 11)

                Case 19: Balance_AumentoHP_Initial = 9 'RandomNumber(7, 11)

                Case 18: Balance_AumentoHP_Initial = 8 'RandomNumber(6, 10)

                Case Else: Balance_AumentoHP_Initial = 7 'RandomNumber(6, UserConstitucion \ 2)
            End Select

        Case eClass.Paladin

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 10 'RandomNumber(9, 11)

                Case 20: Balance_AumentoHP_Initial = 9.5 'RandomNumber(8, 11)

                Case 19: Balance_AumentoHP_Initial = 9 'RandomNumber(7, 11)

                Case 18: Balance_AumentoHP_Initial = 8 'RandomNumber(6, 11)

                Case Else: Balance_AumentoHP_Initial = 7 'RandomNumber(4, UserConstitucion \ 2) + AdicionalHPCazador
            End Select

        Case eClass.Thief

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 7.5 'RandomNumber(6, 9)

                Case 20: Balance_AumentoHP_Initial = 7 'RandomNumber(5, 9)

                Case 19: Balance_AumentoHP_Initial = 6.5 'RandomNumber(4, 9)

                Case 18: Balance_AumentoHP_Initial = 6 'RandomNumber(4, 8)

                Case Else: Balance_AumentoHP_Initial = RandomNumber(4, UserConstitucion \ 2)
            End Select
                       
        Case eClass.Mage

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 7.5 'RandomNumber(6, 9)

                Case 20: Balance_AumentoHP_Initial = 6.5 'RandomNumber(5, 8)

                Case 19: Balance_AumentoHP_Initial = 6 'RandomNumber(4, 8)

                Case 18: Balance_AumentoHP_Initial = 5.5 'RandomNumber(3, 8)

                Case Else: Balance_AumentoHP_Initial = 4 'RandomNumber(5, UserConstitucion \ 2) - AdicionalHPCazador
            End Select
                    
            If Balance_AumentoHP_Initial < 1 Then Balance_AumentoHP_Initial = 4
                    
                    
        Case eClass.Cleric

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 8.5 'RandomNumber(7, 10)

                Case 20: Balance_AumentoHP_Initial = 8 'RandomNumber(6, 10)

                Case 19: Balance_AumentoHP_Initial = 7.5 'RandomNumber(6, 9)

                Case 18: Balance_AumentoHP_Initial = 7 'RandomNumber(5, 9)

                Case Else: Balance_AumentoHP_Initial = 6 'RandomNumber(4, UserConstitucion \ 2)
            End Select
                    
        Case eClass.Druid

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 8.5 'RandomNumber(7, 10)

                Case 20: Balance_AumentoHP_Initial = 8 'RandomNumber(6, 10)

                Case 19: Balance_AumentoHP_Initial = 7.5 'RandomNumber(6, 9)

                Case 18: Balance_AumentoHP_Initial = 7 'RandomNumber(5, 9)

                Case Else: Balance_AumentoHP_Initial = 6 'RandomNumber(4, UserConstitucion \ 2)
            End Select
                     
        Case eClass.Assasin

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 8.5 'RandomNumber(7, 10)

                Case 20: Balance_AumentoHP_Initial = 8 'RandomNumber(6, 10)

                Case 19: Balance_AumentoHP_Initial = 7.5 'RandomNumber(6, 9)

                Case 18: Balance_AumentoHP_Initial = 7 'RandomNumber(5, 9)

                Case Else: Balance_AumentoHP_Initial = 6 'RandomNumber(4, UserConstitucion \ 2)
            End Select

        Case eClass.Bard

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 8.5 'RandomNumber(7, 10)

                Case 20: Balance_AumentoHP_Initial = 8 'RandomNumber(6, 10)

                Case 19: Balance_AumentoHP_Initial = 7.5 'RandomNumber(6, 9)

                Case 18: Balance_AumentoHP_Initial = 7 'RandomNumber(5, 9)

                Case Else: Balance_AumentoHP_Initial = 6 'RandomNumber(4, UserConstitucion \ 2)
            End Select
                    
        Case Else

            Select Case UserConstitucion

                Case 21: Balance_AumentoHP_Initial = 7 'RandomNumber(6, 8)

                Case 20: Balance_AumentoHP_Initial = 6.5 'RandomNumber(5, 8)

                Case 19: Balance_AumentoHP_Initial = 6 'RandomNumber(4, 8)

                Case 18: Balance_AumentoHP_Initial = 5 'RandomNumber(6, 8)

                Case Else: Balance_AumentoHP_Initial = 4 'RandomNumber(5, UserConstitucion \ 2) - AdicionalHPCazador
            End Select
                    
    End Select

    On Error GoTo 0

    Exit Function

Balance_AumentoHP_Initial_Error:

    LogError "Error " & err.Number & " (" & err.Description & ") in procedure Balance_AumentoHP_Initial of Módulo mBalance in line " & Erl
End Function

Public Function Mercader_Range_Armada(ByVal FactionRange As Byte) As String
    Select Case FactionRange
                    
        Case 0
            Mercader_Range_Armada = "<Aprendiz>"
        Case 1
            Mercader_Range_Armada = "<Noble>"
        Case 2
            Mercader_Range_Armada = "<Caballero>"
        Case 3
            Mercader_Range_Armada = "<Capitán>"
        Case 4
            Mercader_Range_Armada = "<Guardián>"
        Case 5
            Mercader_Range_Armada = "<Campeón de la Luz"
    End Select
    
End Function

Public Function Mercader_Range_Legion(ByVal FactionRange As Byte) As String
    Select Case FactionRange
        Case 0
            Mercader_Range_Legion = "<Esbirro>"
        Case 1
            Mercader_Range_Legion = "<Sanguinario>"
        Case 2
            Mercader_Range_Legion = "<Condenado>"
        Case 3
            Mercader_Range_Legion = "<Caballero de la Oscuridad>"
        Case 4
            Mercader_Range_Legion = "<Demonio Infernal>"
        Case 5
            Mercader_Range_Legion = "<Devorador de Almas>"
    End Select
    
End Function
