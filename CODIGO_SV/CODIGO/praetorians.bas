Attribute VB_Name = "PraetoriansCoopNPC"

''**************************************************************
'' PraetoriansCoopNPC.bas - Handles the Praeorians NPCs.
''
'' Implemented by Mariano Barrou (El Oso)
''**************************************************************
'
''**************************************************************************
''This program is free software; you can redistribute it and/or modify
''it under the terms of the Affero General Public License;
''either version 1 of the License, or any later version.
''
''This program is distributed in the hope that it will be useful,
''but WITHOUT ANY WARRANTY; without even the implied warranty of
''MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''Affero General Public License for more details.
''
''You should have received a copy of the Affero General Public License
''along with this program; if not, you can find it at http://www.affero.org/oagpl.html
''**************************************************************************
'
'Option Explicit
''''''''''''''''''''''''''''''''''''''''''
''' DECLARACIONES DEL MODULO PRETORIANO ''
''''''''''''''''''''''''''''''''''''''''''
''' Estas constantes definen que valores tienen
''' los NPCs pretorianos en el NPC-HOSTILES.DAT
''' Son FIJAS, pero se podria hacer una rutina que
''' las lea desde el npcshostiles.dat
'Public Const PRCLER_NPC As Integer = 900   ''"Sacerdote Pretoriano"
'Public Const PRGUER_NPC As Integer = 901   ''"Guerrero  Pretoriano"
'Public Const PRMAGO_NPC As Integer = 902   ''"Mago Pretoriano"
'Public Const PRCAZA_NPC As Integer = 903   ''"Cazador Pretoriano"
'Public Const PRKING_NPC As Integer = 904   ''"Rey Pretoriano"
'
'
'' 1 rey.
'' 3 guerres.
'' 1 caza.
'' 1 mago.
'' 2 clerigos.
'Public Const NRO_PRETORIANOS As Integer = 8
'
'' Contiene los index de los miembros del clan.
'Public Pretorianos(1 To NRO_PRETORIANOS) As Integer
'
'

''''''''''''''''''''''''''''''''''''''''''''''
''Estos numeros son necesarios por cuestiones de
''sonido. Son los numeros de los wavs del cliente.
Public Const SONIDO_DRAGON_VIVO As Integer = 30

'''ALCOBAS REALES
'''OJO LOS BICHOS TAN HARDCODEADOS, NO CAMBIAR EL MAPA DONDE
'''ESTÁN UBICADOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'''MUCHO MENOS LA COORDENADA Y DE LAS ALCOBAS YA QUE DEBE SER LA MISMA!!!
'''(HAY FUNCIONES Q CUENTAN CON QUE ES LA MISMA!)
'Public Const ALCOBA1_X As Integer = 35
'Public Const ALCOBA1_Y As Integer = 25
'Public Const ALCOBA2_X As Integer = 67
'Public Const ALCOBA2_Y As Integer = 25



' Contains all the pretorian's combinations, and its the offsets
Public PretorianAIOffset(1 To 7) As Integer


Public Type tCombinaciones
        NpcIndex() As Integer
        MaxNpc As Integer
        Map As Integer
        X As Integer
        Y As Integer
        
        RespawnTime As Long
        Time As Long
End Type

Public PretorianDatNumbers() As tCombinaciones

'
''Added by Nacho
''Cuantos pretorianos vivos quedan. Uno por cada alcoba
'Public pretorianosVivos As Integer
'

Public Sub Pretorians_Loop()
        '<EhHeader>
        On Error GoTo Pretorians_Loop_Err
        '</EhHeader>
        Dim A As Long
    
100     For A = LBound(PretorianDatNumbers) To UBound(PretorianDatNumbers)
102         If PretorianDatNumbers(A).Time > 0 Then
104             PretorianDatNumbers(A).Time = PretorianDatNumbers(A).Time - 1
            
106             If PretorianDatNumbers(A).Time <= 0 Then
108                 If Not ClanPretoriano(A).SpawnClan(PretorianDatNumbers(A).Map, PretorianDatNumbers(A).X, PretorianDatNumbers(A).Y, A) Then
110                     Call LogError("Error Al cargar boss nro" & A)
                    End If
                End If
            End If
    
112     Next A
        '<EhFooter>
        Exit Sub

Pretorians_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.PraetoriansCoopNPC.Pretorians_Loop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub LoadPretorianData()

        '<EhHeader>
        On Error GoTo LoadPretorianData_Err

        '</EhHeader>

        Dim PretorianDat As String
        
        Dim Manager      As clsIniManager
        
        Set Manager = New clsIniManager
        
100     PretorianDat = DatPath & "Pretorianos.dat"
        
        Manager.Initialize PretorianDat

        Dim NroCombinaciones As Integer

102     NroCombinaciones = val(Manager.GetValue("INIT", "LAST"))

104     ReDim PretorianDatNumbers(0 To NroCombinaciones) As tCombinaciones

        Dim TempInt        As Integer

        Dim Counter        As Long

        Dim PretorianIndex As Integer

        Dim A              As Long, B As Long
        
        ReDim ClanPretoriano(0 To NroCombinaciones) As clsClanPretoriano
        
        Dim Text() As String
        
        Dim Temp As String
        
        For A = 1 To NroCombinaciones
            Set ClanPretoriano(A) = New clsClanPretoriano
            PretorianDatNumbers(A).MaxNpc = val(Manager.GetValue(CStr(A), "MaxNPC"))
            
            Text = Split(Manager.GetValue(CStr(A), "NpcIndex"), "-")
            
            ReDim PretorianDatNumbers(A).NpcIndex(1 To UBound(Text) + 1) As Integer
            
            For B = LBound(Text) To UBound(Text)
                PretorianDatNumbers(A).NpcIndex(B + 1) = val(Text(B))
            Next B
            
            Temp = Manager.GetValue(CStr(A), "Map")
            PretorianDatNumbers(A).Map = val(ReadField(1, Temp, 45))
            PretorianDatNumbers(A).X = val(ReadField(2, Temp, 45))
            PretorianDatNumbers(A).Y = val(ReadField(3, Temp, 45))
            PretorianDatNumbers(A).RespawnTime = val(Manager.GetValue(CStr(A), "RESPAWN"))
        Next A
        
        For A = 1 To NroCombinaciones
            If PretorianDatNumbers(A).Map > 0 Then
                If Not ClanPretoriano(A).SpawnClan(PretorianDatNumbers(A).Map, PretorianDatNumbers(A).X, PretorianDatNumbers(A).Y, A) Then
                    Call LogError("Error Al cargar boss nro" & A)
    
                End If
            End If
        Next A
        
        '<EhFooter>
        Exit Sub

LoadPretorianData_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.PraetoriansCoopNPC.LoadPretorianData " & "at line " & Erl

        

        '</EhFooter>
End Sub

