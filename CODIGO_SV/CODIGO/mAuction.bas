Attribute VB_Name = "mAuction"
Option Explicit

Public Const AUCTION_TIME As Long = 300

Public Type tAuction_Security
    IP As String
    Email As String
End Type

Public Type tAuction_Offer
    Name As String
    Gld As Long
    Eldhir As Long
    TimeLastOffer As Integer
    Security As tAuction_Security
End Type

Public Type tAuction
    Name As String
    
    ObjIndex As Integer
    Amount As Integer
    Durabilidad As Integer
    Gld As Long
    Eldhir As Long
    Time As Long
    
    Offer As tAuction_Offer
    
    Security As tAuction_Security
End Type


Public Auction As tAuction

Private Function Auction_Checking_SameUser(ByVal UserIndex As Integer, _
                                      ByVal Email As String, _
                                      ByVal IP As String) As Boolean
        '<EhHeader>
        On Error GoTo Auction_Checking_SameUser_Err
        '</EhHeader>
    
        Dim Account As String
    
        ' Mismo Email
100     If StrComp(UserList(UserIndex).Account.Email, Email) = 0 Then
            Exit Function
        End If
    
    
        ' Misma IP
102     If UserList(UserIndex).IpAddress <> vbNullString Then
104         If StrComp(UserList(UserIndex).IpAddress, IP) = 0 Then
                Exit Function
            End If
        End If
    
106     Auction_Checking_SameUser = True
        '<EhFooter>
        Exit Function

Auction_Checking_SameUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_Checking_SameUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Auction_CreateNew(ByVal UserIndex As Integer, _
                             ByVal ObjIndex As Integer, _
                             ByVal Amount As Integer, _
                             ByVal Gld As Long, _
                             ByVal Eldhir As Long)
        '<EhHeader>
        On Error GoTo Auction_CreateNew_Err
        '</EhHeader>
                             
            If UserList(UserIndex).Account.Premium < 2 Then
                     Call WriteConsoleMsg(UserIndex, "Debes ser al menos Tier 2 para poder subastar objetos. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
        
        
100     If UserList(UserIndex).Pos.Map <> 1 Then
102         Call WriteConsoleMsg(UserIndex, "No puedes subastar objetos si no estas en la ciudad principal.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
        
        
    
104     If Auction.ObjIndex > 0 Then
106         Call WriteConsoleMsg(UserIndex, "Ya hay una subasta en trámite, espera a que termine.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
108     If ConfigServer.ModoSubastas = 0 Then
110         Call WriteConsoleMsg(UserIndex, "Las subastas no estan permitidas momentaneamente.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
112     If Not Auction_ObjValid(ObjIndex) Then
114         Call WriteConsoleMsg(UserIndex, "No puedes subastar este objeto.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
116     With Auction
118         .Name = UCase$(UserList(UserIndex).Name)
120         .ObjIndex = ObjIndex
122         .Amount = Amount
124         .Gld = Gld
126         .Eldhir = Eldhir
128         .Time = AUCTION_TIME
        
130         .Security.Email = UserList(UserIndex).Account.Email
132         .Security.IP = UserList(UserIndex).IpAddress
134         .Offer.Name = UCase$(UserList(UserIndex).Name)
136         .Offer.Gld = Gld
138         .Offer.Eldhir = Eldhir
        End With
    
    
        Dim TempObj As String
140     TempObj = ObjData(ObjIndex).Name
    
142     If ObjData(ObjIndex).Bronce = 1 Then
144         TempObj = TempObj & " [BRONCE]"
        End If
    
146     If ObjData(ObjIndex).Plata = 1 Then
148         TempObj = TempObj & " [PLATA]"
        End If
    
150     If ObjData(ObjIndex).Oro = 1 Then
152         TempObj = TempObj & " [ORO]"
        End If
    
154     If ObjData(ObjIndex).Premium = 1 Then
156         TempObj = TempObj & " [PREMIUM]"
        End If
    
158     Call QuitarObjetos(ObjIndex, Amount, UserIndex)
160     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» Un objeto se está vendiendo al mejor postor y es " & TempObj & " (x" & Amount & ")", FontTypeNames.FONTTYPE_CRITICO))
162     Call Logs_Security(eSecurity, eSubastas, "Subasta nueva» El personaje " & Auction.Offer.Name & " puso el objeto " & TempObj & " (x" & Amount & ") a " & Format$(Gld, "#,###") & " Monedas de Oro Y " & Eldhir & " Monedas de Eldhir.")
        '<EhFooter>
        Exit Sub

Auction_CreateNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_CreateNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Auction_Offer(ByVal UserIndex As Integer, _
                         ByVal Gld As Long, _
                         ByVal Eldhir As Long)
        '<EhHeader>
        On Error GoTo Auction_Offer_Err
        '</EhHeader>
    
        Dim tUser As Integer
        Dim FilePath As String
        Dim TempGld As Long
        Dim TempEldhir As Long
    
100     If ConfigServer.ModoSubastas = 0 Then
102         Call WriteConsoleMsg(UserIndex, "Las subastas no estan permitidas momentaneamente.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
104     If Auction.ObjIndex = 0 Then
106         Call WriteConsoleMsg(UserIndex, "¡No hay ninguna subasta en trámite!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
108     If Auction.Name = UCase$(UserList(UserIndex).Name) Then
110         Call WriteConsoleMsg(UserIndex, "¡No te ofrezcas a ti mismo!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
112     If Not Auction_Checking_SameUser(UserIndex, Auction.Security.Email, Auction.Security.IP) Then
114         Call Logs_Security(eSecurity, eAntiHack, "[DOBLE CLIENTE] El personaje " & UCase$(UserList(UserIndex).Name) & " con IP: " & Auction.Security.IP & " Y Email: " & Auction.Security.Email & " intentó ofrecerse a sí mismo y fue advertido.")
116         Call WriteConsoleMsg(UserIndex, "¡No te ofrezcas a ti mismo!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
118     If Not Auction_Checking_SameUser(UserIndex, Auction.Offer.Security.Email, Auction.Offer.Security.IP) Then
120         Call Logs_Security(eSecurity, eAntiHack, "[DOBLE CLIENTE] El personaje " & UCase$(UserList(UserIndex).Name) & " con IP: " & Auction.Offer.Security.IP & " Y Email: " & Auction.Offer.Security.Email & " intentó ofrecerse a sí mismo y fue advertido.")
122         Call WriteConsoleMsg(UserIndex, "¡La última oferta la has realizado tu!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    

    
124     If Auction.Offer.Name = UCase$(UserList(UserIndex).Name) Then
126         Call WriteConsoleMsg(UserIndex, "¡Ya has ofrecido!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
128     If UserList(UserIndex).Stats.Gld < Gld Then
            Exit Sub
        End If
    
130     If UserList(UserIndex).Stats.Eldhir < Eldhir Then
            Exit Sub
        End If
    

    
132     With Auction.Offer
134         If .TimeLastOffer > 0 Then
136             Call WriteConsoleMsg(UserIndex, "Debes esperar unos momentos para realizar otra oferta.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
    
138         If Gld < (.Gld * 1.1) Then
140             Call WriteConsoleMsg(UserIndex, "¡Debes ofreer al menos un 10% más que la oferta anterior. En total sería (" & .Gld * 1.1 & "!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub
            End If
    
    
            ' Le reintegramos el oro a la anterior oferta.
142         If .Name <> vbNullString And .Name <> Auction.Name Then
144             tUser = NameIndex(.Name)
            
146             If tUser > 0 Then
148                 UserList(tUser).Stats.Gld = UserList(tUser).Stats.Gld + .Gld
150                 UserList(tUser).Stats.Eldhir = UserList(tUser).Stats.Eldhir + .Eldhir
                
152                 Call WriteUpdateUserStats(tUser)
154                 Call WriteConsoleMsg(tUser, "¡Han ofrecido más Monedas que tú! ¿Te darás por vencido?", FontTypeNames.FONTTYPE_INFORED)
                Else
156                 FilePath = CharPath & .Name & ".chr"
158                 TempGld = val(GetVar(FilePath, "STATS", "GLD"))
160                 TempEldhir = val(GetVar(FilePath, "STATS", "ELDHIR"))
                
162                 Call WriteVar(FilePath, "STATS", "GLD", CStr(TempGld + .Gld))
164                 Call WriteVar(FilePath, "STATS", "ELDHIR", CStr(TempEldhir + .Eldhir))
                End If
            End If
        
166         .Name = UCase$(UserList(UserIndex).Name)
168         .Gld = Gld
170         .Eldhir = Eldhir
172         .TimeLastOffer = 5
174         .Security.Email = UserList(UserIndex).Account.Email
176         .Security.IP = UserList(UserIndex).IpAddress
        
178         UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - Gld
180         UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir - Eldhir
182         Call WriteUpdateUserStats(UserIndex)
        
184         Call SendData(SendTarget.toMap, 1, PrepareMessageConsoleMsg("Subasta» El personaje " & Auction.Offer.Name & " ha ofrecido: " & Format$(.Gld, "#,###") & " Monedas de Oro Y " & .Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_INFOGREEN))
186         Call Logs_Security(eSecurity, eSubastas, "El personaje " & Auction.Offer.Name & " ha ofrecido: " & .Gld & " Monedas de Oro Y " & .Eldhir & " Monedas de Eldhir.")
        End With
    
188     Auction.Time = Auction.Time + 20
        '<EhFooter>
        Exit Sub

Auction_Offer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_Offer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Auction_Loop()
        '<EhHeader>
        On Error GoTo Auction_Loop_Err
        '</EhHeader>
100     With Auction
102         If .Offer.TimeLastOffer > 0 Then
104             .Offer.TimeLastOffer = .Offer.TimeLastOffer - 1
            End If
        
106         If .Time > 0 Then
108             .Time = .Time - 1

110             If (.Time Mod 60) = 0 Then
                
112                 If .Time = 0 Then
114                     If .Offer.Name = .Name Then
116                         Call Auction_RewardObj
118                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» La subasta del objeto " & ObjData(.ObjIndex).Name & " (x" & .Amount & ") ha concluído sin ofertas.", FontTypeNames.FONTTYPE_CRITICO))
                        
                        Else
                    
120                         Call Auction_RewardGld
122                         Call Auction_RewardObj

124                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» La subasta del objeto " & ObjData(.ObjIndex).Name & " (x" & .Amount & ") ha concluído. ¡Se lo ha llevado el personaje " & .Offer.Name & "!", FontTypeNames.FONTTYPE_CRITICO))
                        
                        End If
                    
126                     Call Auction_Reset
128                 ElseIf .Time <> 60 Then
130                     If .Offer.Name <> .Name Then
132                         Call SendData(SendTarget.toMap, 1, PrepareMessageConsoleMsg("Subasta» " & ObjData(.ObjIndex).Name & " (x" & .Amount & "). La última oferta es de " & Format$(.Offer.Gld, "#,###") & " Monedas de Oro y " & .Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_CRITICO))
                        End If
                    Else
134                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» Última posibilidad para ofertar por el objeto " & ObjData(.ObjIndex).Name & " (x" & .Amount & "). La última oferta es de " & Format$(.Offer.Gld, "#,###") & " Monedas de Oro y " & .Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_CRITICO))
                    End If
                End If
            End If
    
        End With
        '<EhFooter>
        Exit Sub

Auction_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_Loop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Auction_Object()
        '<EhHeader>
        On Error GoTo Auction_Object_Err
        '</EhHeader>
        Dim A As Long
        Dim FilePath As String
        Dim Temp As String
    
100     FilePath = CharPath & Auction.Offer.Name & ".chr"
    
102     For A = 1 To MAX_BANCOINVENTORY_SLOTS
104         Temp = GetVar(FilePath, "BANCOINVENTORY", "OBJ" & A)
        
106         If Temp = "0-0" Then
108             Call WriteVar(FilePath, "BANCOINVENTORY", "OBJ" & A, Auction.ObjIndex & "-" & Auction.Amount)
110             Call Logs_Security(eSecurity, eSubastas, "El personaje " & Auction.Offer.Name & " recibió en su boveda: " & ObjData(Auction.ObjIndex).Name & " (x" & Auction.Amount & ")")
                Exit Sub
            End If
112     Next A
    
114     Call Logs_Security(eSecurity, eSubastas, "El personaje " & Auction.Offer.Name & " no recibió el objeto en su boveda: " & ObjData(Auction.ObjIndex).Name & " (x" & Auction.Amount & ")")
        '<EhFooter>
        Exit Sub

Auction_Object_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_Object " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Auction_RewardGld()
        '<EhHeader>
        On Error GoTo Auction_RewardGld_Err
        '</EhHeader>
        Dim tUser As Integer
        Dim FilePath As String
        Dim TempGld As Long
        Dim TempEldhir As Long
    
100     tUser = NameIndex(Auction.Name)
        
        ' Subastador
102     If tUser > 0 Then
104         Call WriteConsoleMsg(tUser, "Has recibido el dinero de tu subasta. ¡Has acumulado " & Auction.Offer.Gld & " Monedas de Oro Y " & Auction.Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_INFOGREEN)
106         UserList(tUser).Stats.Gld = UserList(tUser).Stats.Gld + Auction.Offer.Gld
108         UserList(tUser).Stats.Eldhir = UserList(tUser).Stats.Eldhir + Auction.Offer.Eldhir
110         Call WriteUpdateUserStats(tUser)
        Else
112         FilePath = CharPath & Auction.Name & ".chr"
114         TempGld = val(GetVar(FilePath, "STATS", "GLD"))
116         TempEldhir = val(GetVar(FilePath, "STATS", "ELDHIR"))
            
118         Call WriteVar(FilePath, "STATS", "GLD", CStr(TempGld + Auction.Offer.Gld))
120         Call WriteVar(FilePath, "STATS", "ELDHIR", CStr(TempEldhir + Auction.Offer.Eldhir))
        End If
        '<EhFooter>
        Exit Sub

Auction_RewardGld_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_RewardGld " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Auction_RewardObj()
        '<EhHeader>
        On Error GoTo Auction_RewardObj_Err
        '</EhHeader>

        Dim tUser As Integer
        Dim FilePath As String
        Dim TempGld As Long
        Dim TempEldhir As Long
    
        Dim Obj As Obj
    
100     With Auction
            ' Personaje que recibe el objeto
102         tUser = NameIndex(.Offer.Name)
            
104         If tUser > 0 Then
106             Obj.Amount = Auction.Amount
108             Obj.ObjIndex = Auction.ObjIndex
            
110             If Not MeterItemEnInventario(tUser, Obj) Then
112                 Call Logs_Security(eSecurity, eSubastas, "El personaje " & .Offer.Name & " no tenia espacio en inventario, por lo que no recibió el objeto de la subasta: " & ObjData(.ObjIndex).Name & " (x" & .Amount & ")")
                Else
114                 Call Logs_Security(eSecurity, eSubastas, "El personaje " & .Offer.Name & " recibió " & ObjData(.ObjIndex).Name & " (x" & .Amount & ") ofertando " & .Offer.Gld & " Monedas de Oro Y " & .Offer.Eldhir & " Monedas de Eldhir.")
                End If
            Else
116             Call Auction_Object
            End If
        End With
        '<EhFooter>
        Exit Sub

Auction_RewardObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_RewardObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub Auction_Reset()
        '<EhHeader>
        On Error GoTo Auction_Reset_Err
        '</EhHeader>

        Dim NullAuction As tAuction
    
100     Auction = NullAuction

        '<EhFooter>
        Exit Sub

Auction_Reset_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_Reset " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Function Auction_ObjValid(ByVal ObjIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Auction_ObjValid_Err
        '</EhHeader>
100     If ObjData(ObjIndex).OBJType = otBebidas Or _
            ObjData(ObjIndex).OBJType = otBotellaLlena Or _
            ObjData(ObjIndex).OBJType = otBotellaVacia Or _
            ObjData(ObjIndex).OBJType = otPociones Or _
            ObjData(ObjIndex).OBJType = otUseOnce Then
        
            Exit Function
        End If
    
102     Auction_ObjValid = True
        '<EhFooter>
        Exit Function

Auction_ObjValid_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_ObjValid " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Auction_Cancel()
        '<EhHeader>
        On Error GoTo Auction_Cancel_Err
        '</EhHeader>
100     If Auction.ObjIndex = 0 Then Exit Sub
102     Call Auction_RewardObj
104     Call Auction_Reset
        '<EhFooter>
        Exit Sub

Auction_Cancel_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAuction.Auction_Cancel " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
