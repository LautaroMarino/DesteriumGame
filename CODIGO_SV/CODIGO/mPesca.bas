Attribute VB_Name = "mPesca"
' Sistema de pesca de objetos

Public Type tPesca

    ObjIndex As Integer
    Amount As Integer
    Probability As Byte

End Type

Public Pesca_NumItems As Byte

Public PescaItem()    As tPesca

Public Sub Pesca_LoadItems()
        '<EhHeader>
        On Error GoTo Pesca_LoadItems_Err
        '</EhHeader>

        Dim Manager As clsIniManager

        Dim A       As Long

        Dim Temp    As String
    
100     Set Manager = New clsIniManager
    
102     Manager.Initialize DatPath & "PESCA.DAT"
    
104     Pesca_NumItems = val(Manager.GetValue("INIT", "ITEMS"))
    
106     ReDim PescaItem(1 To Pesca_NumItems) As tPesca

108     For A = 1 To Pesca_NumItems
110         Temp = Manager.GetValue("INIT", A)
        
112         With PescaItem(A)
114             .ObjIndex = val(ReadField(1, Temp, 45))
116             .Amount = val(ReadField(2, Temp, 45))
118             .Probability = val(ReadField(3, Temp, 45))
            End With

120     Next A
    
122     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

Pesca_LoadItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mPesca.Pesca_LoadItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Pesca_ExtractItem(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Pesca_ExtractItem_Err
        '</EhHeader>

        Dim A           As Long, B As Long

        Dim RandomItem  As Byte, Random As Byte

        Dim Probability As Byte

        Dim Obj         As Obj
    
100     For A = 1 To Pesca_NumItems
        
102         With PescaItem(A)

104             For B = 1 To .Probability

                    ' 10% de ir pasando de etapas
106                 If RandomNumber(1, 100) <= 10 Then
108                     Probability = Probability + 1
                    Else

                        Exit For

                    End If

110             Next B
    
                ' Si cumplimos con la etapa requerida:
112             If Probability = .Probability Then
114                 Obj.ObjIndex = .ObjIndex
116                 Obj.Amount = .Amount
                            
118                 If Not MeterItemEnInventario(UserIndex, Obj) Then
120                     Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj)
                    End If
                
122                 Call WriteConsoleMsg(UserIndex, "Has recolectado de las profundidades del mar " & ObjData(.ObjIndex).Name & " (x" & .Amount & ")", FontTypeNames.FONTTYPE_INFO)
                    'Else
                    ' Call WriteConsoleMsg(UserIndex, Probability & "/" & .Probability, FontTypeNames.FONTTYPE_INFO)
                End If
            
124             Probability = 0
            End With

126     Next A

        '<EhFooter>
        Exit Sub

Pesca_ExtractItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mPesca.Pesca_ExtractItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
