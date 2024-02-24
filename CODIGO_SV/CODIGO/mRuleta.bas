Attribute VB_Name = "mRuleta"
Option Explicit

' Sistema de Ruleta

Public Type tRuletaItem

    ObjIndex As Integer      ' Index del objeto
    Amount As Integer       ' Cantidad de objeto que da
    Prob As Byte                '1,2,3,4,5
    ProbNum As Byte         '10,20,30,40,50,60,70,80,90a99

End Type

Public Type tRuletaConfig
    ItemLast As Integer
    Items() As tRuletaItem
    RuletaGld As Long
    RuletaDsp As Long
    

End Type

Public RuletaConfig As tRuletaConfig

Public Sub Ruleta_LoadItems()

    Dim Manager As clsIniManager

    Dim A       As Long

    Dim Temp    As String
    
    Dim FilePath As String
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "ruleta.dat"
    
    Manager.Initialize FilePath
    
    With RuletaConfig
        .ItemLast = val(Manager.GetValue("INIT", "LAST"))
        .RuletaDsp = val(Manager.GetValue("INIT", "RULETADSP"))
        .RuletaGld = val(Manager.GetValue("INIT", "RULETAGLD"))
        
        If .ItemLast > 0 Then
            ReDim .Items(1 To .ItemLast) As tRuletaItem
        
            For A = 1 To .ItemLast

                With .Items(A)
                    Temp = Manager.GetValue("LIST", "OBJ" & A)
                
                    .ObjIndex = val(ReadField(1, Temp, 45))
                    .Amount = val(ReadField(2, Temp, 45))
                    .Prob = val(ReadField(3, Temp, 45))
                    .ProbNum = val(ReadField(4, Temp, 45))
                
                End With

            Next A
    
        End If
    
    End With
    
    Manager.DumpFile DatPath & "client\ruleta.dat"
    Set Manager = Nothing

End Sub

Public Sub Ruleta_Tirada(ByVal UserIndex As Integer, ByVal Mode As Byte)
        '<EhHeader>
        On Error GoTo Ruleta_Tirada_Err
        '</EhHeader>

100     With UserList(UserIndex)

            Exit Sub

102         If Mode = 1 Then ' Monedas de Oro
104             If .Stats.Gld < RuletaConfig.RuletaGld Then
106                 Call WriteConsoleMsg(UserIndex, "No tienes suficientes Monedas de Oro.", FontTypeNames.FONTTYPE_INFORED)
                    'TODO: Enter task description here
                    Exit Sub

                End If
            
108             .Stats.Gld = .Stats.Gld - RuletaConfig.RuletaGld
110             Call WriteUpdateGold(UserIndex)
112         ElseIf Mode = 2 Then            ' Monedas DSP

114             If .Stats.Eldhir < RuletaConfig.RuletaDsp Then
116                 Call WriteConsoleMsg(UserIndex, "No tienes suficientes DSP.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If
            
118             .Stats.Eldhir = .Stats.Eldhir - RuletaConfig.RuletaDsp
120             Call WriteUpdateDsp(UserIndex)
            End If
        
122         Call Ruleta_Tirada_Item(UserIndex)
        End With

        '<EhFooter>
        Exit Sub

Ruleta_Tirada_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRuleta.Ruleta_Tirada " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Ruleta_Tirada_Item(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Ruleta_Tirada_Item_Err
        '</EhHeader>
    
        Dim RandItem    As Byte

        Dim A           As Long, S As Long
    
        Dim Probability As Long, Sound As Long
    
        Dim MiObj As Obj
    
100     RandItem = RandomNumber(1, RuletaConfig.ItemLast)
    
102     With RuletaConfig.Items(RandItem)

104         For A = 1 To .Prob

106             If RandomNumber(1, 100) <= .ProbNum Then
108                 Probability = Probability + 1

                End If

110         Next A
                
112         If Probability = .Prob Then
114             MiObj.Amount = .Amount
116             MiObj.ObjIndex = .ObjIndex

118             If Not MeterItemEnInventario(UserIndex, MiObj, True) Then
120                 Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

                End If
        
                Exit Sub
            
122             S = RandomNumber(1, 100)
        
124             If S <= 25 Then
126                 Sound = eSound.sChestDrop1
128             ElseIf S <= 50 Then
130                 Sound = eSound.sChestDrop2
                Else
132                 Sound = eSound.sChestDrop3

                End If
        
134             Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(Sound, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

            End If
    
        End With

        '<EhFooter>
        Exit Sub

Ruleta_Tirada_Item_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mRuleta.Ruleta_Tirada_Item " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
