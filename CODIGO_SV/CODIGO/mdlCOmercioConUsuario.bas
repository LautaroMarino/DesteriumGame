Attribute VB_Name = "mdlCOmercioConUsuario"
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000

Private Const MAX_OBJ_LOGUEABLE As Long = 100

Public Const MAX_OFFER_SLOTS    As Integer = 20

Public Const GOLD_OFFER_SLOT    As Integer = MAX_OFFER_SLOTS + 1

Public Const ELDHIR_OFFER_SLOT  As Integer = MAX_OFFER_SLOTS + 2

Public Type tCOmercioUsuario

    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto(1 To MAX_OFFER_SLOTS) As Integer 'Indice de los objetos que se desea dar
    GoldAmount As Long
    EldhirAmount As Long
    
    cant(1 To MAX_OFFER_SLOTS) As Long 'Cuantos objetos desea dar
    Acepto As Boolean
    Confirmo As Boolean

End Type

Private Type tOfferItem

    ObjIndex As Integer
    Amount As Long

End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)

    '***************************************************
    'Autor: Unkown
    'Last Modification: 25/11/2009
    '
    '***************************************************
    On Error GoTo ErrHandler
    
    'Si ambos pusieron /comerciar entonces
    If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
       
        If UserList(Origen).flags.Comerciando Or UserList(Destino).flags.Comerciando Then
            Call WriteConsoleMsg(Origen, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(Destino, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)

            Exit Sub

        End If
        
        'Actualiza el inventario del usuario
        Call UpdateUserInv(True, Origen, 0)
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Origen)
        UserList(Origen).flags.Comerciando = True
    
        'Actualiza el inventario del usuario
        Call UpdateUserInv(True, Destino, 0)
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Destino)
        UserList(Destino).flags.Comerciando = True
        
        Call Logs_User(UserList(Origen).Name, eLog.eUser, eLogDescUser.eOther, "Inicio de comercio con " & UserList(Destino).Name & " Cuenta: " & UserList(Destino).Account.Email)
        Call Logs_User(UserList(Destino).Name, eLog.eUser, eLogDescUser.eOther, "Inicio de comercio con " & UserList(Origen).Name & " Cuenta: " & UserList(Origen).Account.Email)
        
        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        Call WriteConsoleMsg(Origen, "Le has ofrecido comerciar al personaje " & UserList(Destino).Name & ".", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen
        
    End If
    
    Call FlushBuffer(Destino)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

Public Sub EnviarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte)
        '<EhHeader>
        On Error GoTo EnviarOferta_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unkown
        'Last Modification: 25/11/2009
        'Sends the offer change to the other trading user
        '25/11/2009: ZaMa - Implementado nuevo sistema de comercio con ofertas variables.
        '***************************************************
        Dim ObjIndex       As Integer

        Dim ObjAmount      As Long

        Dim OtherUserIndex As Integer
    
100     OtherUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
102     With UserList(OtherUserIndex)

104         If OfferSlot = GOLD_OFFER_SLOT Then
106             ObjIndex = iORO
108             ObjAmount = .ComUsu.GoldAmount
110         ElseIf OfferSlot = ELDHIR_OFFER_SLOT Then
112             ObjIndex = iELDHIR
114             ObjAmount = .ComUsu.EldhirAmount
            Else
116             ObjIndex = .ComUsu.Objeto(OfferSlot)
118             ObjAmount = .ComUsu.cant(OfferSlot)
            End If

        End With
   
120     Call WriteChangeUserTradeSlot(UserIndex, OfferSlot, ObjIndex, ObjAmount)

        '<EhFooter>
        Exit Sub

EnviarOferta_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mdlCOmercioConUsuario.EnviarOferta " & _
               "at line " & Erl & " Cuenta: " & UserList(UserIndex).Account.Email
        
        '</EhFooter>
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo FinComerciarUsu_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unkown
        'Last Modification: 25/11/2009
        '25/11/2009: ZaMa - Limpio los arrays (por el nuevo sistema)
        '***************************************************
        Dim i As Long
    
100     With UserList(UserIndex)

102         If .ComUsu.DestUsu > 0 Then
104             Call WriteUserCommerceEnd(UserIndex)
            End If
        
106         .ComUsu.Acepto = False
108         .ComUsu.Confirmo = False
110         .ComUsu.DestUsu = 0
        
112         For i = 1 To MAX_OFFER_SLOTS
114             .ComUsu.cant(i) = 0
116             .ComUsu.Objeto(i) = 0
118         Next i
        
120         .ComUsu.EldhirAmount = 0
122         .ComUsu.GoldAmount = 0
124         .ComUsu.DestNick = vbNullString
126         .flags.Comerciando = False
        End With

        '<EhFooter>
        Exit Sub

FinComerciarUsu_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mdlCOmercioConUsuario.FinComerciarUsu " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo AceptarComercioUsu_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unkown
        'Last Modification: 06/05/2010
        '25/11/2009: ZaMa - Ahora se traspasan hasta 5 items + oro al comerciar
        '06/05/2010: ZaMa - Ahora valida si los usuarios tienen los items que ofertan.
        '***************************************************
        Dim TradingObj    As Obj

        Dim OtroUserIndex As Integer

        Dim OfferSlot     As Integer

100     UserList(UserIndex).ComUsu.Acepto = True
    
102     OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
        ' Acepto el otro?
104     If UserList(OtroUserIndex).ComUsu.Acepto = False Then

            Exit Sub

        End If
    
        ' User valido?
106     If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
108         Call FinComerciarUsu(UserIndex)

            Exit Sub

        End If
    
        ' Aceptaron ambos, chequeo que tengan los items que ofertaron
110     If Not HasOfferedItems(UserIndex) Then
        
112         Call WriteConsoleMsg(UserIndex, "¡¡¡El comercio se canceló porque no posees los ítems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
114         Call WriteConsoleMsg(OtroUserIndex, "¡¡¡El comercio se canceló porque " & UserList(UserIndex).Name & " no posee los ítems que ofertó!!!", FontTypeNames.FONTTYPE_FIGHT)
        
116         Call FinComerciarUsu(UserIndex)
118         Call FinComerciarUsu(OtroUserIndex)
120         Call Protocol.FlushBuffer(OtroUserIndex)
        
            Exit Sub
        
122     ElseIf Not HasOfferedItems(OtroUserIndex) Then
        
124         Call WriteConsoleMsg(UserIndex, "¡¡¡El comercio se canceló porque " & UserList(OtroUserIndex).Name & " no posee los ítems que ofertó!!!", FontTypeNames.FONTTYPE_FIGHT)
126         Call WriteConsoleMsg(OtroUserIndex, "¡¡¡El comercio se canceló porque no posees los ítems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
        
128         Call FinComerciarUsu(UserIndex)
130         Call FinComerciarUsu(OtroUserIndex)
132         Call Protocol.FlushBuffer(OtroUserIndex)
        
            Exit Sub
        
        End If
    
        ' Envio los items a quien corresponde
134     For OfferSlot = 1 To MAX_OFFER_SLOTS + 2
        
            ' Items del 1er usuario
136         With UserList(UserIndex)

                ' Le pasa el oro
138             If OfferSlot = GOLD_OFFER_SLOT Then
                    ' Quito la cantidad de oro ofrecida
140                 .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount

                    ' Log
142                 If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call Logs_User(.Name, eLog.eUser, eDropGld, .Name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                    ' Update Usuario
144                 Call WriteUpdateUserStats(UserIndex)
                    ' Se la doy al otro
146                 UserList(OtroUserIndex).Stats.Gld = UserList(OtroUserIndex).Stats.Gld + .ComUsu.GoldAmount
                    ' Update Otro Usuario
148                 Call WriteUpdateUserStats(OtroUserIndex)
            
150             ElseIf OfferSlot = ELDHIR_OFFER_SLOT Then
                    ' Quito la cantidad de oro ofrecida
152                 .Stats.Eldhir = .Stats.Eldhir - .ComUsu.EldhirAmount

                    ' Log
154                 If .ComUsu.EldhirAmount > MAX_ORO_LOGUEABLE Then Call Logs_User(.Name, eLog.eUser, eDropEldhir, .Name & " soltó Eldhir en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.EldhirAmount)
                    ' Update Usuario
156                 Call WriteUpdateUserStats(UserIndex)
                    ' Se la doy al otro
158                 UserList(OtroUserIndex).Stats.Eldhir = UserList(OtroUserIndex).Stats.Eldhir + .ComUsu.EldhirAmount
                    ' Update Otro Usuario
160                 Call WriteUpdateUserStats(OtroUserIndex)
                    ' Le pasa lo ofertado de los slots con items
162             ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
164                 TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
166                 TradingObj.Amount = .ComUsu.cant(OfferSlot)
                                
                    'Quita el objeto y se lo da al otro
168                 If Not MeterItemEnInventario(OtroUserIndex, TradingObj) Then
170                     Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj)
                    End If
            
172                 Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, UserIndex)
                
174                 If ObjData(TradingObj.ObjIndex).OBJType = otTransformVIP Then
176                     If UserList(UserIndex).flags.TransformVIP = 1 Then
178                         Call TransformVIP_User(UserIndex, 0)
                        End If
                    End If
                
                    'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
180                 If ObjData(TradingObj.ObjIndex).Log = 1 Then
182                     Call Logs_User(.Name, eLog.eUser, eCommerce_Obj, .Name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
            
                    'Es mucha cantidad?
184                 If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then

                        'Si no es de los prohibidos de loguear, lo logueamos.
186                     If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
188                         Call Logs_User(UserList(OtroUserIndex).Name, eLog.eUser, eCommerce_Obj, UserList(OtroUserIndex).Name & " le pasó en comercio seguro a " & .Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                        End If
                    End If
                End If

            End With
        
            ' Items del 2do usuario
190         With UserList(OtroUserIndex)

                ' Le pasa el oro
192             If OfferSlot = GOLD_OFFER_SLOT Then
                    ' Quito la cantidad de oro ofrecida
194                 .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                    ' Log
                
196                 If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call Logs_User(.Name, eLog.eUser, eDropGld, .Name & " soltó oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                    ' Update Usuario
198                 Call WriteUpdateUserStats(OtroUserIndex)
                    'y se la doy al otro
200                 UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + .ComUsu.GoldAmount
                
202                 If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eDropGld, UserList(UserIndex).Name & " recibió oro en comercio seguro con " & .Name & ". Cantidad: " & .ComUsu.GoldAmount)
                    ' Update Otro Usuario
204                 Call WriteUpdateUserStats(UserIndex)
                
206             ElseIf OfferSlot = ELDHIR_OFFER_SLOT Then
                    ' Quito la cantidad de oro ofrecida
208                 .Stats.Eldhir = .Stats.Eldhir - .ComUsu.EldhirAmount
                    ' Log
                
210                 If .ComUsu.EldhirAmount > MAX_ORO_LOGUEABLE Then Call Logs_User(.Name, eLog.eUser, eDropEldhir, .Name & " soltó Eldhir en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & .ComUsu.EldhirAmount)
                    ' Update Usuario
212                 Call WriteUpdateUserStats(OtroUserIndex)
                    'y se la doy al otro
214                 UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir + .ComUsu.EldhirAmount
                
216                 If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eDropEldhir, UserList(UserIndex).Name & " recibió Eldhir en comercio seguro con " & .Name & ". Cantidad: " & .ComUsu.EldhirAmount)
                    ' Update Otro Usuario
218                 Call WriteUpdateUserStats(UserIndex)
                    ' Le pasa la oferta de los slots con items
220             ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
222                 TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
224                 TradingObj.Amount = .ComUsu.cant(OfferSlot)
                                
                    'Quita el objeto y se lo da al otro
226                 If Not MeterItemEnInventario(UserIndex, TradingObj) Then
228                     Call TirarItemAlPiso(UserList(UserIndex).Pos, TradingObj)
                    End If
            
230                 Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, OtroUserIndex)
                
232                 If ObjData(TradingObj.ObjIndex).OBJType = otTransformVIP Then
234                     If UserList(OtroUserIndex).flags.TransformVIP = 1 Then
236                         Call TransformVIP_User(OtroUserIndex, 0)
                        End If
                    End If
        
                    'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
238                 If ObjData(TradingObj.ObjIndex).Log = 1 Then
240                     Call Logs_User(.Name, eLog.eUser, eCommerce_Obj, .Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
            
                    'Es mucha cantidad?
242                 If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then

                        'Si no es de los prohibidos de loguear, lo logueamos.
244                     If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
246                         Call Logs_User(.Name, eLog.eUser, eCommerce_Obj, .Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                        End If
                    End If
                End If

            End With
        
248     Next OfferSlot

        ' End Trade
250     Call FinComerciarUsu(UserIndex)
252     Call FinComerciarUsu(OtroUserIndex)
254     Call Protocol.FlushBuffer(OtroUserIndex)
    
        '<EhFooter>
        Exit Sub

AceptarComercioUsu_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mdlCOmercioConUsuario.AceptarComercioUsu " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AgregarOferta(ByVal UserIndex As Integer, _
                         ByVal OfferSlot As Byte, _
                         ByVal ObjIndex As Integer, _
                         ByVal Amount As Long, _
                         ByVal IsGold As Boolean, _
                         ByVal IsEldhir As Boolean)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 24/11/2009
        'Adds gold or items to the user's offer
        '***************************************************
        '<EhHeader>
        On Error GoTo AgregarOferta_Err
        '</EhHeader>

100     If PuedeSeguirComerciando(UserIndex) Then

102         With UserList(UserIndex).ComUsu

                ' Si ya confirmo su oferta, no puede cambiarla!
104             If Not .Confirmo Then
106                 If IsGold Then
                        ' Agregamos (o quitamos) mas oro a la oferta
108                     .GoldAmount = .GoldAmount + Amount
                    
                        ' Imposible que pase, pero por las dudas..
110                     If .GoldAmount < 0 Then .GoldAmount = 0
112                 ElseIf IsEldhir Then
                        ' Agregamos (o quitamos) mas oro a la oferta
114                     .EldhirAmount = .EldhirAmount + Amount
                    
                        ' Imposible que pase, pero por las dudas..
116                     If .EldhirAmount < 0 Then .EldhirAmount = 0
                
                    Else

                        ' Agreamos (o quitamos) el item y su cantidad en el slot correspondiente
                        ' Si es 0 estoy modificando la cantidad, no agregando
118                     If ObjIndex > 0 Then .Objeto(OfferSlot) = ObjIndex
120                     .cant(OfferSlot) = .cant(OfferSlot) + Amount
                    
                        'Quitó todos los items de ese tipo
122                     If .cant(OfferSlot) <= 0 Then
                            ' Removemos el objeto para evitar conflictos
124                         .Objeto(OfferSlot) = 0
126                         .cant(OfferSlot) = 0
                        End If
                    End If
                End If

            End With

        End If

        '<EhFooter>
        Exit Sub

AgregarOferta_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mdlCOmercioConUsuario.AgregarOferta " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function PuedeSeguirComerciando(ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo PuedeSeguirComerciando_Err
        '</EhHeader>

        '***************************************************
        'Autor: ZaMa
        'Last Modification: 24/11/2009
        'Validates wether the conditions for the commerce to keep going are satisfied
        '***************************************************
        Dim OtroUserIndex    As Integer

        Dim ComercioInvalido As Boolean

100     With UserList(UserIndex)

            ' Usuario valido?
102         If .ComUsu.DestUsu <= 0 Or .ComUsu.DestUsu > MaxUsers Then
104             ComercioInvalido = True
            End If
    
106         OtroUserIndex = .ComUsu.DestUsu
    
108         If Not ComercioInvalido Then

                ' Estan logueados?
110             If UserList(OtroUserIndex).flags.UserLogged = False Or .flags.UserLogged = False Then
112                 ComercioInvalido = True
                End If
            End If
    
114         If Not ComercioInvalido Then

                ' Se estan comerciando el uno al otro?
116             If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
118                 ComercioInvalido = True
                End If
            End If
    
120         If Not ComercioInvalido Then

                ' El nombre del otro es el mismo que al que le comercio?
122             If UserList(OtroUserIndex).Name <> .ComUsu.DestNick Then
124                 ComercioInvalido = True
                End If
            End If
    
126         If Not ComercioInvalido Then

                ' Mi nombre  es el mismo que al que el le comercia?
128             If .Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
130                 ComercioInvalido = True
                End If
            End If
    
132         If Not ComercioInvalido Then

                ' Esta vivo?
134             If UserList(OtroUserIndex).flags.Muerto = 1 Then
136                 ComercioInvalido = True
                End If
            End If
    
            ' Fin del comercio
138         If ComercioInvalido = True Then
140             Call FinComerciarUsu(UserIndex)
        
142             If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
144                 Call FinComerciarUsu(OtroUserIndex)
146                 Call Protocol.FlushBuffer(OtroUserIndex)
                End If
        
                Exit Function

            End If

        End With

148     PuedeSeguirComerciando = True

        '<EhFooter>
        Exit Function

PuedeSeguirComerciando_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mdlCOmercioConUsuario.PuedeSeguirComerciando " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function HasOfferedItems(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 05/06/2010
        'Checks whether the user has the offered items in his inventory or not.
        '***************************************************
        '<EhHeader>
        On Error GoTo HasOfferedItems_Err
        '</EhHeader>

        Dim OfferedItems(MAX_OFFER_SLOTS - 1) As tOfferItem

        Dim Slot                              As Long

        Dim SlotAux                           As Long

        Dim SlotCount                         As Long
    
        Dim ObjIndex                          As Integer
    
100     With UserList(UserIndex).ComUsu
        
            ' Agrupo los items que son iguales
102         For Slot = 1 To MAX_OFFER_SLOTS
                    
104             ObjIndex = .Objeto(Slot)
            
106             If ObjIndex > 0 Then
            
108                 For SlotAux = 0 To SlotCount - 1
                    
110                     If ObjIndex = OfferedItems(SlotAux).ObjIndex Then
                            ' Son iguales, aumento la cantidad
112                         OfferedItems(SlotAux).Amount = OfferedItems(SlotAux).Amount + .cant(Slot)

                            Exit For

                        End If
                    
114                 Next SlotAux
                
                    ' No encontro otro igual, lo agrego
116                 If SlotAux = SlotCount Then
118                     OfferedItems(SlotCount).ObjIndex = ObjIndex
120                     OfferedItems(SlotCount).Amount = .cant(Slot)
                    
122                     SlotCount = SlotCount + 1
                    End If
                
                End If
            
124         Next Slot
        
            ' Chequeo que tengan la cantidad en el inventario
126         For Slot = 0 To SlotCount - 1

128             If Not HasEnoughItems(UserIndex, OfferedItems(Slot).ObjIndex, OfferedItems(Slot).Amount) Then Exit Function
130         Next Slot
        
            ' Compruebo que tenga el oro que oferta
132         If UserList(UserIndex).Stats.Gld < .GoldAmount Then Exit Function
        
            ' Compruebo que tenga el Eldhir que oferta
134         If UserList(UserIndex).Stats.Eldhir < .EldhirAmount Then Exit Function
        End With
    
136     HasOfferedItems = True

        '<EhFooter>
        Exit Function

HasOfferedItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mdlCOmercioConUsuario.HasOfferedItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
