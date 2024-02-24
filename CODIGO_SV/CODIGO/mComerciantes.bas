Attribute VB_Name = "mComerciantes"
Option Explicit

' Sistema de comerciantes alquilables por los usuarios. Para que puedan tener sus propios comerciantes y vender sus items. Incluso items de donación al valor que vale el item en el juego.
' ARCHIVO COMERCIANTES.DAT ¡SERVIDOR!

Public Type tCommerceChar
    Char As String
End Type

Public Type tComerciantes
    Owner As String
    OwnerDate As String
    
    ValueGLD As Long
    ValueDSP As Long
    
    
    RewardGLD As Long           ' ORO que gano vendiendo
    RewardDSP As Long           ' DSP que gano vendiendo
    
    
    MaxItems As Byte
    Pos As WorldPos
    NpcIndex As Integer
    
    
    Days As Double
    
    BkItems As Inventario
    
End Type

Public ComerciantesLast As Integer
Public Comerciantes() As tComerciantes
Public CommerceChar As tCommerceChar

Public Sub Comerciantes_Load()
        '<EhHeader>
        On Error GoTo Comerciantes_Load_Err
        '</EhHeader>
    
        Dim Manager  As clsIniManager

        Dim FilePath As String

        Dim A        As Long, B As Long

        Dim Temp     As String

        Dim NpcIndex As Integer
    
100     Set Manager = New clsIniManager
    
102     FilePath = DatPath & "comerciantes.dat"
    
104     Manager.Initialize FilePath
    
106     ComerciantesLast = val(Manager.GetValue("INIT", "LAST"))
            
            If ComerciantesLast > 0 Then
108     ReDim Comerciantes(1 To ComerciantesLast) As tComerciantes
    
110     For A = 1 To ComerciantesLast

112         With Comerciantes(A)
114             .MaxItems = val(Manager.GetValue(A, "MAXITEMS"))
116             .ValueDSP = val(Manager.GetValue(A, "VALUEDSP"))
118             .ValueGLD = val(Manager.GetValue(A, "VALUEGLD"))
120             .NpcIndex = val(Manager.GetValue(A, "NPCINDEX"))
122             .Owner = Manager.GetValue(A, "OWNER")
124             .OwnerDate = Manager.GetValue(A, "OWNERDATE")
126             .Days = val(Manager.GetValue(A, "DAYS"))
            
128             Temp = Manager.GetValue(A, "POSITION")
            
130             .Pos.Map = val(ReadField(1, Temp, 45))
132             .Pos.X = val(ReadField(2, Temp, 45))
134             .Pos.Y = val(ReadField(3, Temp, 45))
            
136             NpcIndex = SpawnNpc(.NpcIndex, .Pos, False, False)
            
138             If NpcIndex = 0 Then
140                 Call LogError("ERROR CRITICO EN LA CARGA DE COMERCIANTES")
142                 Set Manager = Nothing
                    Exit Sub
                Else
144                 .NpcIndex = NpcIndex
146                 Npclist(NpcIndex).CommerceIndex = A
148                 Npclist(NpcIndex).CommerceChar = .Owner
                
                    ' Le carga el inventario que tenía !
150                 If Npclist(NpcIndex).Invent.NroItems > 0 Then
152                     Call Manager.ChangeValue(A, "INVENTORY_CANT", CStr(Npclist(NpcIndex).Invent.NroItems))
                
154                     For B = 1 To Npclist(NpcIndex).Invent.NroItems
156                         Temp = Manager.GetValue(A, "INVENTORY_OBJ" & B)
                        
158                         Npclist(NpcIndex).Invent.Object(B).ObjIndex = val(ReadField(1, Temp, 45))
160                         Npclist(NpcIndex).Invent.Object(B).Amount = val(ReadField(2, Temp, 45))
162                     Next B
                
                    End If

                End If
            
            End With
    
164     Next A

End If

    
166     Set Manager = Nothing

        '<EhFooter>
        Exit Sub

Comerciantes_Load_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mComerciantes.Comerciantes_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' Guarda los comerciantes actualizados con el nombre del dueño del comerciante
Public Sub Comerciantes_Save()
        '<EhHeader>
        On Error GoTo Comerciantes_Save_Err
        '</EhHeader>

        Dim Manager  As clsIniManager

        Dim FilePath As String
        Dim A As Long
        Dim B As Long
    
100     FilePath = DatPath & "Comerciantes.dat"
    
102     Set Manager = New clsIniManager
    
104     For A = 1 To ComerciantesLast

106         With Comerciantes(A)
108             Call Manager.ChangeValue(A, "OWNER", .Owner)
110             Call Manager.ChangeValue(A, "OWNERDATE", .OwnerDate)

                ' Guarda lo que haya obtenido
                 Call Manager.ChangeValue(A, "REWARDDSP", CStr(.RewardDSP))
                Call Manager.ChangeValue(A, "REWARDGLD", CStr(.RewardGLD))
                  
                ' Guarda todos los objetos que tenga en ese momento.
112             If Npclist(.NpcIndex).Invent.NroItems > 0 Then
114                 Call Manager.ChangeValue(A, "INVENTORY_CANT", CStr(Npclist(.NpcIndex).Invent.NroItems))
                
116                 For B = 1 To Npclist(.NpcIndex).Invent.NroItems
118                     Call Manager.ChangeValue(A, "INVENTORY_OBJ" & B, CStr(Npclist(.NpcIndex).Invent.Object(B).ObjIndex) & "-" & CStr(Npclist(.NpcIndex).Invent.Object(B).Amount))
120                 Next B
                
                End If
            End With

122     Next A
    
124     Manager.DumpFile FilePath
    
126     Set Manager = Nothing

        '<EhFooter>
        Exit Sub

Comerciantes_Save_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mComerciantes.Comerciantes_Save " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' Comprueba las fechas de los comerciantes. Si tiene que cancelar le devuelve los objetos al pibe.
' Se comprueba cada un minuto ya que no tiene que ser tan preciso. [CAMBIAR DESDE LA LLAMADA]
Public Sub Comerciantes_Loop()
        '<EhHeader>
        On Error GoTo Comerciantes_Loop_Err
        '</EhHeader>

        Dim A As Long
        Dim NullComerciante As tComerciantes
    
100     For A = 1 To ComerciantesLast

102         With Comerciantes(A)

104             If .Owner <> vbNullString Then
106                 If Format(Now, "dd/mm/yyyy") > .OwnerDate Then
108                     Call Comerciantes_Return_User(A)
                    End If

                End If

            End With

110     Next A

        '<EhFooter>
        Exit Sub

Comerciantes_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mComerciantes.Comerciantes_Loop " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' Devuelve los objetos a la boveda del personaje.
' En caso de que supere los slots disponibles de bovedas, se pierden los items, pero aun asi se los registramos para no ser tan bruscos
Public Sub Comerciantes_Return_User(ByVal ComercianteIndex As Integer)
        '<EhHeader>
        On Error GoTo Comerciantes_Return_User_Err
        '</EhHeader>

        Dim UserName As String
        Dim tUser As Integer
        Dim FilePath As String
    
100     With Comerciantes(ComercianteIndex)
102         UserName = .Owner
        
104         tUser = NameIndex(UserName)
        
106         If tUser > 0 Then
108             Call WriteConsoleMsg(tUser, "Te hemos devuelto los objetos que han quedado y están en tu boveda...", FontTypeNames.FONTTYPE_INFOGREEN)
        
            Else
110             FilePath = CharPath & UserName & ".chr"
            
            End If
        
112         .Owner = vbNullString
114         .OwnerDate = vbNullString
        
116         Call Comerciantes_Save
        End With
        '<EhFooter>
        Exit Sub

Comerciantes_Return_User_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mComerciantes.Comerciantes_Return_User " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' Comprueba que el pibe pueda alquilar la tienda
Public Function Commerce_CanUser_Rent(ByVal UserIndex As Integer, _
                                      ByRef Commerce As tComerciantes) As Boolean
        '<EhHeader>
        On Error GoTo Commerce_CanUser_Rent_Err
        '</EhHeader>

100     With UserList(UserIndex)

102         If .Stats.Gld < Commerce.ValueGLD Then
104             Call WriteConsoleMsg(UserIndex, "¡No tienes suficientes Monedas de Oro para alquilar esta tienda.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function

            End If
    
106         If .Stats.Eldhir < Commerce.ValueDSP Then
108             Call WriteConsoleMsg(UserIndex, "¡No tienes suficientes Monedas Desterium para alquilar esta tienda. Recuerda tenerlas en tu billetera y no en tu cuenta.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function

            End If
            
            If Commerce.Owner <> vbNullString Then
                 Call WriteConsoleMsg(UserIndex, "¡El mercader está alquilado y estará disponible el " & Commerce.OwnerDate & ".", FontTypeNames.FONTTYPE_INFORED)
            
                Exit Function
            End If

        End With
    
110     Commerce_CanUser_Rent = True

        '<EhFooter>
        Exit Function

Commerce_CanUser_Rent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mComerciantes.Commerce_CanUser_Rent " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub Commerce_ViewBalance(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    
    Dim Temp As tComerciantes

    Temp = Comerciantes(Npclist(NpcIndex).CommerceIndex)
    
    Call WriteChatOverHead(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro", Npclist(NpcIndex).Char.charindex, vbCyan)
    Call WriteConsoleMsg(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro", FontTypeNames.FONTTYPE_INFOGREEN)

End Sub

Public Sub Commerce_ReclamarGanancias(ByVal NpcIndex As Integer, _
                                      ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Commerce_ReclamarGanancias_Err
        '</EhHeader>
    
        Dim CI   As Integer
    
        Dim Temp As tComerciantes

100     CI = Npclist(NpcIndex).CommerceIndex
102     Temp = Comerciantes(CI)
    
104     Call WriteChatOverHead(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro Y ¡Has retirado TODO!", Npclist(NpcIndex).Char.charindex, vbCyan)
106     Call WriteConsoleMsg(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro Y ¡Has retirado TODO!", FontTypeNames.FONTTYPE_INFOGREEN)
    
108     With UserList(UserIndex)
110         .Stats.Gld = .Stats.Gld + Temp.RewardGLD
112         .Stats.Eldhir = .Stats.Eldhir + Temp.RewardDSP
        
114         Call WriteUpdateUserStats(UserIndex)
        End With
    
116     With Comerciantes(CI)
118         .RewardDSP = 0
120         .RewardGLD = 0
    
        End With

        '<EhFooter>
        Exit Sub

Commerce_ReclamarGanancias_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mComerciantes.Commerce_ReclamarGanancias " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' El comerciante alquila un nuevo comerciante
Public Sub Commerce_SetNew(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    
    Dim Slot As Byte

    Dim Temp As tComerciantes
    
    Slot = Npclist(NpcIndex).CommerceIndex
    Temp = Comerciantes(Slot)
    
    If Not Commerce_CanUser_Rent(UserIndex, Temp) Then Exit Sub
    
    With Comerciantes(Slot)
        .Owner = UCase$(UserList(UserIndex).Name)
        .OwnerDate = DateAdd("d", .Days, Now)
    
         Npclist(NpcIndex).CommerceChar = .Owner
    End With
    
    With UserList(UserIndex)
        .Stats.Gld = .Stats.Gld - Temp.ValueGLD
        .Stats.Eldhir = .Stats.Eldhir - Temp.ValueDSP
    End With
    
    Call WriteConsoleMsg(UserIndex, "¡El cielo no tiene limites! Has alquilado el mercado hasta el día " & Comerciantes(Slot).OwnerDate & ". Esperamos que puedas vender todos tus objetos", FontTypeNames.FONTTYPE_USERPREMIUM)
    Call Logs_Security(eLog.eSecurity, eLogSecurity.eComerciantes, "Alquiler del comerciante " & Npclist(NpcIndex).Name & ".")
    Call WriteUpdateUserStats(UserIndex)
    
End Sub


' DESHABILITADO
Public Sub Commerce_Can_AddItem(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer)
    
    Dim CommerceIndex As Integer
    
    CommerceIndex = Npclist(NpcIndex).CommerceIndex
    
    With Comerciantes(CommerceIndex)
        If Npclist(NpcIndex).Invent.NroItems = .MaxItems Then
            'Call writeconsolemsg(Userindex,"¡Este comerciante admite solo " & A & " espacios para vender."
        
        
        End If
    
    End With
End Sub
