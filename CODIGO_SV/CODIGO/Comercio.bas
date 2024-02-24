Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio

    Compra = 1
    Venta = 2

End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

' Chequeamos que exista un mismo objeto para poder venderlo aquí.
Private Function Comercio_CheckItem(ByVal NpcIndex As Integer, _
                                    ByVal ObjIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Comercio_CheckItem_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_INVENTORY_SLOTS

102         With Npclist(NpcIndex).Invent

104             If .Object(A).ObjIndex = ObjIndex Then
106                 Comercio_CheckItem = True

                    Exit Function

                End If

            End With

108     Next A

        '<EhFooter>
        Exit Function

Comercio_CheckItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.Comercio_CheckItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, _
                    ByVal UserIndex As Integer, _
                    ByVal NpcIndex As Integer, _
                    ByVal Slot As Integer, _
                    ByVal cantidad As Integer, _
                    ByVal SelectedPrice As Byte)

        '<EhHeader>
        On Error GoTo Comercio_Err

        '</EhHeader>

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 07/06/2010
        '27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
        '  - 06/13/08 (NicoNZ)
        '07/06/2010: ZaMa - Los objetos se loguean si superan la cantidad de 1k (antes era solo si eran 1k).
        '*************************************************
        Dim PrecioDiamanteRojo As Long

        Dim PrecioDiamanteAzul As Long
            
        Dim PrecioPoints As Long
        
        Dim Objeto             As Obj
    
100     If cantidad < 1 Or Slot < 1 Then Exit Sub
102     If SelectedPrice > 1 Then Exit Sub
          
104     If Modo = eModoComercio.Compra Then
106         If Slot > MAX_INVENTORY_SLOTS Then

                Exit Sub

108         ElseIf cantidad > MAX_INVENTORY_OBJS Then
110             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
112             Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & cantidad)
114             UserList(UserIndex).flags.Ban = 1
            
116             Call Protocol.Kick(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
 
                Exit Sub

118         ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then

                Exit Sub

            End If
            
120         'Objeto.Amount = cantidad
122         Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
              
124         If ObjData(Objeto.ObjIndex).Upgrade.RequiredCant = 0 And ObjData(Objeto.ObjIndex).Points = 0 Then
126             If SelectedPrice = 0 And ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).valor = 0 Then Exit Sub
128             If SelectedPrice = 1 And ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).ValorEldhir = 0 Then Exit Sub

            End If
              
130         If cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(Slot).Amount
            
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
              
132         If SelectedPrice = 0 Then
134             PrecioDiamanteRojo = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).valor / Descuento(UserIndex) * cantidad) + 0.5)
                
136             If UserList(UserIndex).Stats.Gld < PrecioDiamanteRojo Then
138                 Call WriteConsoleMsg(UserIndex, "No tienes suficientes Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

140         ElseIf SelectedPrice = 1 Then
142             PrecioDiamanteAzul = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).ValorEldhir / Descuento(UserIndex) * cantidad) + 0.5)

144             If UserList(UserIndex).Stats.Eldhir < PrecioDiamanteAzul Then
146                 Call WriteConsoleMsg(UserIndex, "No tienes suficientes DSP.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
        
            End If
            
            PrecioPoints = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Points / Descuento(UserIndex) * cantidad) + 0.5)

            If UserList(UserIndex).Stats.Points < PrecioPoints Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes Puntos Desterium.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            Dim SlotEvent As Byte

148         SlotEvent = UserList(UserIndex).flags.SlotEvent
        
150         If SlotEvent > 0 Then
152             If Events(SlotEvent).LimitRed > 0 Then
154                 If Objeto.ObjIndex = POCION_ROJA Then
156                     Call WriteConsoleMsg(UserIndex, "No puedes comprar pociones rojas en éste tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If

                End If
            
158             If Events(SlotEvent).ChangeClass > 0 Or Events(SlotEvent).ChangeRaze > 0 Or Events(SlotEvent).ChangeLevel > 0 Then
160                 If Events(SlotEvent).TimeCancel > 0 Then
162                     Call WriteConsoleMsg(UserIndex, "Aún no está habilitada la compra de objetos. Espera a que se completen los cupos.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If

                End If

            End If
                        
            Dim A As Long
            
            Objeto.Amount = cantidad
            
            ' @ Comprueba si tiene los objetos necesarios
164         If ObjData(Objeto.ObjIndex).Upgrade.RequiredCant > 0 Then
166             '  cantidad = 1        ' salvo las flechas es 1
                    
                If ObjData(Objeto.ObjIndex).OBJType = otFlechas Then
                    Objeto.Amount = 500
                    cantidad = 500
                Else
                    Objeto.Amount = 1
                    cantidad = 1

                End If
                    
168             For A = 1 To ObjData(Objeto.ObjIndex).Upgrade.RequiredCant

170                 If Not TieneObjetos(ObjData(Objeto.ObjIndex).Upgrade.Required(A).ObjIndex, ObjData(Objeto.ObjIndex).Upgrade.Required(A).Amount, UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "No tienes " & ObjData(Objeto.ObjIndex).Upgrade.Required(A).ObjIndex & " (x " & ObjData(Objeto.ObjIndex).Upgrade.Required(A).Amount & ")", FontTypeNames.FONTTYPE_INFORED)
                        Exit Sub

                    End If

172             Next A
                
174             For A = 1 To ObjData(Objeto.ObjIndex).Upgrade.RequiredCant
176                 Call QuitarObjetos(ObjData(Objeto.ObjIndex).Upgrade.Required(A).ObjIndex, ObjData(Objeto.ObjIndex).Upgrade.Required(A).Amount, UserIndex)
178             Next A

            End If
        
180         If MeterItemEnInventario(UserIndex, Objeto) = False Then Exit Sub
            
            ' # Compra un objeto por duración. Lo asignamos en un array()
            If ObjData(Objeto.ObjIndex).DurationDay > 0 Then
                Dim TempDate As String
                TempDate = DateAdd("d", ObjData(Objeto.ObjIndex).DurationDay, Now)
                Call Bonus_AddUser_Online(UserIndex, eBonusType.eObj, Objeto.ObjIndex, Objeto.Amount, 0, TempDate, False)
            End If
            
182         If SelectedPrice = 0 Then
184             UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - PrecioDiamanteRojo
186         ElseIf SelectedPrice = 1 Then
188             UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir - PrecioDiamanteAzul

            End If
            
            UserList(UserIndex).Stats.Points = UserList(UserIndex).Stats.Points - PrecioPoints
            
            Dim CI As Integer
            
            CI = Npclist(NpcIndex).CommerceIndex

            If CI > 0 Then
                Comerciantes(CI).RewardDSP = Comerciantes(CI).RewardDSP + PrecioDiamanteAzul
                Comerciantes(CI).RewardGLD = Comerciantes(CI).RewardGLD + PrecioDiamanteRojo

            End If
        
190         Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), cantidad)
        
192         If ObjData(Objeto.ObjIndex).Log = 1 Then
194             Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBuyObj, "compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
196         ElseIf Objeto.Amount >= 100 Then 'Es mucha cantidad?

                'Si no es de los prohibidos de loguear, lo logueamos.
198             If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
200                 Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBuyObj, "compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)

                End If

            End If
        
            'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
202         If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
204             Call WriteVar(Npcs_FilePath, "NPC" & Npclist(NpcIndex).numero, "obj" & Slot, Objeto.ObjIndex & "-0")
206             Call logVentaCasa(UserList(UserIndex).Name & " compró " & ObjData(Objeto.ObjIndex).Name)

            End If
        
208     ElseIf Modo = eModoComercio.Venta Then

210         If cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
        
212         Objeto.Amount = cantidad
214         Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
216         If Objeto.ObjIndex = 0 Then

                Exit Sub
256         ElseIf UserList(UserIndex).Invent.Object(Slot).Amount < 0 Or cantidad = 0 Then

                Exit Sub

258         ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then

                Exit Sub
                
            ElseIf Npclist(NpcIndex).NPCtype = eCommerceChar And Npclist(NpcIndex).CommerceChar <> UCase$(UserList(UserIndex).Name) Then
                Call WriteConsoleMsg(UserIndex, "Sólo el dueño del personaje puede agregar objetos al mercado. ¿Deseas alquilarme luego? /ALQUILAR", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
218         ElseIf ObjData(Objeto.ObjIndex).OBJType = otMonturas Then
220             Call WriteConsoleMsg(UserIndex, "No puedes vender tu montura.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
222         ElseIf ObjData(Objeto.ObjIndex).NoNada = 1 Then
224             Call WriteConsoleMsg(UserIndex, "No puedes realizar ninguna acción con este objeto. ¡Podría ser de uso personal!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub
226         ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then

228             If Npclist(NpcIndex).Name <> "SR" Then
230                 Call WriteConsoleMsg(UserIndex, "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

232         ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

234             If Npclist(NpcIndex).Name <> "SC" Then
236                 Call WriteConsoleMsg(UserIndex, "Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

238         ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType) Or Objeto.ObjIndex = iORO Then

240             If Npclist(NpcIndex).TipoItems = 999 Then
242                 If ObjData(Objeto.ObjIndex).ValorEldhir <= 0 Then
244                     Call WriteConsoleMsg(UserIndex, "¡Ja ja ja! Vende tus baratijas en aquel mercado.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                End If

246             If Npclist(NpcIndex).TipoItems <> 1000 Then
                
                    ' Criaturas que venden items de todos los tipos.
248                 If Npclist(NpcIndex).TipoItems = 998 Then
250                     If Not Comercio_CheckItem(NpcIndex, Objeto.ObjIndex) Then
252                         Call WriteConsoleMsg(UserIndex, "Lo siento, debes vender tus objetos en el mercado global, no aquí.", FontTypeNames.FONTTYPE_INFO)
        
                            Exit Sub

                        End If

                    Else
254                     Call WriteConsoleMsg(UserIndex, "Lo siento, debes vender tus objetos en el mercado global, no aquí.", FontTypeNames.FONTTYPE_INFO)
    
                        Exit Sub

                    End If

                End If

260         ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.SemiDios Then
262             Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
        
264         If ObjData(Objeto.ObjIndex).OBJType = otGemaTelep Then
266             Call WriteConsoleMsg(UserIndex, "No puedes vender este objeto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
                
268         If ObjData(Objeto.ObjIndex).OBJType = otTransformVIP Then
270             Call WriteConsoleMsg(UserIndex, "No puedes vender este objeto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
272         If SelectedPrice = 1 Then
274             Call WriteConsoleMsg(UserIndex, "Lo siento Joven, deberás vender tu objeto por Monedas de Oro o bien a los usuarios del Servidor.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If

            If Npclist(NpcIndex).NPCtype <> eNPCType.eCommerceChar Then
                ' Comprueba si no tenía que vender el Objeto para pasar a la próxima misión.
276             Call Quests_AddSale(UserIndex, Objeto.ObjIndex, Objeto.Amount)
278             Call QuitarUserInvItem(UserIndex, Slot, cantidad)
        
280             PrecioDiamanteRojo = Fix(SalePrice(Objeto.ObjIndex) * cantidad)
282             PrecioDiamanteAzul = Fix(SalePriceDiamanteAzul(Objeto.ObjIndex) * cantidad)
        
284             If SelectedPrice = 0 Then
286                 UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + PrecioDiamanteRojo
288             ElseIf SelectedPrice = 1 Then
290                 UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir + PrecioDiamanteAzul

                End If

            Else

                If ObjData(Objeto.ObjIndex).ValorEldhir = 0 And ObjData(Objeto.ObjIndex).valor = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Por el momento estos objetos no puedes venderlos aquí. Pero espera pronta noticias", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If

                Call QuitarUserInvItem(UserIndex, Slot, cantidad)

            End If
        
292         If UserList(UserIndex).Stats.Gld > MAXORO Then UserList(UserIndex).Stats.Gld = MAXORO
        
294         If UserList(UserIndex).Stats.Eldhir > MAXORO Then UserList(UserIndex).Stats.Eldhir = MAXORO
            
296         If Not (ObjData(Objeto.ObjIndex).LvlMax > 0 And ObjData(Objeto.ObjIndex).LvlMax < UserList(UserIndex).Stats.Elv) Then

                Dim NpcSlot As Integer

298             NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
            
300             If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                    'Mete el obj en el slot
302                 Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
304                 Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount
    
306                 If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
308                     Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS

                    End If
                
310                 Call EnviarNpcInv(NpcSlot, UserIndex, UserList(UserIndex).flags.TargetNPC)

                End If

            End If
            
312         If ObjData(Objeto.ObjIndex).Log = 1 Then
314             Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eSaleObj, "vendió del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name & " al NPC: " & Npclist(NpcIndex).numero)
316         ElseIf Objeto.Amount >= 100 Then 'Es mucha cantidad?

                'Si no es de los prohibidos de loguear, lo logueamos.
318             If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
320                 Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eSaleObj, "vendió del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name & " al NPC: " & Npclist(NpcIndex).numero)

                End If

            End If
        
        End If
    
322     Call UpdateUserInv(False, UserIndex, Slot)
324     Call WriteUpdateGold(UserIndex)
326     Call WriteUpdateDsp(UserIndex)
328     Call WriteTradeOK(UserIndex)
    
332     Call SubirSkill(UserIndex, eSkill.Comerciar, True)
    
        '<EhFooter>
        Exit Sub

Comercio_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.Comercio " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        '<EhHeader>
        On Error GoTo IniciarComercioNPC_Err
        '</EhHeader>
100     Call EnviarNpcInv(0, UserIndex, UserList(UserIndex).flags.TargetNPC)
102     UserList(UserIndex).flags.Comerciando = True
104     Call WriteCommerceInit(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Name, _
                                            Npclist(UserList(UserIndex).flags.TargetNPC).Quest, _
                                            Npclist(UserList(UserIndex).flags.TargetNPC).Quests)
        '<EhFooter>
        Exit Sub

IniciarComercioNPC_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.IniciarComercioNPC " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, _
                              ByVal Objeto As Integer, _
                              ByVal cantidad As Integer) As Integer
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        '<EhHeader>
        On Error GoTo SlotEnNPCInv_Err
        '</EhHeader>
100     SlotEnNPCInv = 1

102     Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + cantidad <= MAX_INVENTORY_OBJS
        
104         SlotEnNPCInv = SlotEnNPCInv + 1

106         If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
        Loop
    
108     If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
110         SlotEnNPCInv = 1
        
112         Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
114             SlotEnNPCInv = SlotEnNPCInv + 1

116             If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
            Loop
        
118         If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    
        End If
    
        '<EhFooter>
        Exit Function

SlotEnNPCInv_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.SlotEnNPCInv " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UpdateSlot As Byte, _
                         ByVal UserIndex As Integer, _
                         ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo EnviarNpcInv_Err
        '</EhHeader>

        '*************************************************
        'Author: Nacho (Integer)
        'Last Modified: 07/08/2022
        ' Actualiza solo los Slots necesarios
        '*************************************************
        Dim Slot As Byte

        Dim val  As Single
    
        Dim val2 As Single
    
        Dim thisObj As Obj
        Dim DummyObj As Obj
    
100     If NpcIndex = 0 Then Exit Sub
    
102     If UpdateSlot > 0 Then
104         If Npclist(NpcIndex).Invent.Object(UpdateSlot).ObjIndex > 0 Then
                
106             thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(UpdateSlot).ObjIndex
108             thisObj.Amount = Npclist(NpcIndex).Invent.Object(UpdateSlot).Amount
                
110             val = (ObjData(thisObj.ObjIndex).valor) / Descuento(UserIndex)
112             val2 = (ObjData(thisObj.ObjIndex).ValorEldhir) / Descuento(UserIndex)
                
114             Call WriteChangeNPCInventorySlot(UserIndex, UpdateSlot, thisObj, val, val2)
            Else
    
            
    
116             Call WriteChangeNPCInventorySlot(UserIndex, UpdateSlot, DummyObj, 0, 0)
            End If
        Else

118         For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS
    
120         If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then

122             thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
124             thisObj.Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount
                
126             val = (ObjData(thisObj.ObjIndex).valor) / Descuento(UserIndex)
128             val2 = (ObjData(thisObj.ObjIndex).ValorEldhir) / Descuento(UserIndex)
                
130             Call WriteChangeNPCInventorySlot(UserIndex, Slot, thisObj, val, val2)
            Else
    
132             Call WriteChangeNPCInventorySlot(UserIndex, Slot, DummyObj, 0, 0)
            End If
        
134     Next Slot

    End If
    
        '<EhFooter>
        Exit Sub

EnviarNpcInv_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.EnviarNpcInv " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePrice(ByVal ObjIndex As Integer) As Single
        '<EhHeader>
        On Error GoTo SalePrice_Err
        '</EhHeader>

        '*************************************************
        'Author: Nicolás (NicoNZ)
        '
        '*************************************************
100     If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
102     If ItemNewbie(ObjIndex) Then Exit Function
    
104     SalePrice = ObjData(ObjIndex).valor / REDUCTOR_PRECIOVENTA
        '<EhFooter>
        Exit Function

SalePrice_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.SalePrice " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePriceDiamanteAzul(ByVal ObjIndex As Integer) As Single
        '<EhHeader>
        On Error GoTo SalePriceDiamanteAzul_Err
        '</EhHeader>

        '*************************************************
        'Author: Nicolás (NicoNZ)
        '
        '*************************************************
100     If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
102     If ItemNewbie(ObjIndex) Then Exit Function
    
104     SalePriceDiamanteAzul = ObjData(ObjIndex).ValorEldhir / REDUCTOR_PRECIOVENTA
        '<EhFooter>
        Exit Function

SalePriceDiamanteAzul_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.SalePriceDiamanteAzul " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        '<EhHeader>
        On Error GoTo Descuento_Err
        '</EhHeader>
100     Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
        '<EhFooter>
        Exit Function

Descuento_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modSistemaComercio.Descuento " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
