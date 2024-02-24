Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
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

Option Explicit

Public Enum E_BANK

    e_User = 1
    e_Account = 2

End Enum

Sub IniciarDeposito(ByVal UserIndex As Integer, ByVal TypeBank As E_BANK)

        '<EhHeader>
        On Error GoTo IniciarDeposito_Err

        '</EhHeader>
                    
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
    
100     Select Case TypeBank

            Case E_BANK.e_User
102             Call UpdateBanUserInv(True, UserIndex, 0)
            
104         Case E_BANK.e_Account

                If UserList(UserIndex).Account.Premium < 2 Then
                    Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior poseen un banco exclusivo. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If

106             Call UpdateBanUserInv_Account(True, UserIndex, 0)
        
        End Select
    
108     Call WriteBankInit(UserIndex, TypeBank)
        'Call WriteUpdateUserStats(UserIndex)
    
110     UserList(UserIndex).flags.Comerciando = True

        '<EhFooter>
        Exit Sub

IniciarDeposito_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.IniciarDeposito " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub SendBanObj(ByVal UserIndex As Integer, _
               ByVal Slot As Byte, _
               ByRef Object As UserOBJ, _
               ByVal TypeBank As E_BANK)
        '<EhHeader>
        On Error GoTo SendBanObj_Err
        '</EhHeader>
               
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
    
100     Select Case TypeBank

            Case E_BANK.e_User
102             UserList(UserIndex).BancoInvent.Object(Slot) = Object
104             Call WriteChangeBankSlot(UserIndex, Slot)
            
106         Case E_BANK.e_Account
108             UserList(UserIndex).Account.BancoInvent.Object(Slot) = Object
110             Call WriteChangeBankSlot_Account(UserIndex, Slot)
        End Select

        '<EhFooter>
        Exit Sub

SendBanObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.SendBanObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, _
                     ByVal UserIndex As Integer, _
                     ByVal Slot As Byte)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UpdateBanUserInv_Err
        '</EhHeader>

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

100     With UserList(UserIndex)

            'Actualiza un solo slot
102         If Not UpdateAll Then

                'Actualiza el inventario
104             If .BancoInvent.Object(Slot).ObjIndex > 0 Then
106                 Call SendBanObj(UserIndex, Slot, .BancoInvent.Object(Slot), e_User)
                Else
108                 Call SendBanObj(UserIndex, Slot, NullObj, e_User)
                End If

            Else

                'Actualiza todos los slots
110             For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                    'Actualiza el inventario
112                 If .BancoInvent.Object(LoopC).ObjIndex > 0 Then
114                     Call SendBanObj(UserIndex, LoopC, .BancoInvent.Object(LoopC), e_User)
                    Else
116                     Call SendBanObj(UserIndex, LoopC, NullObj, e_User)
                    End If
            
118             Next LoopC

            End If

        End With

        '<EhFooter>
        Exit Sub

UpdateBanUserInv_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UpdateBanUserInv " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UpdateBanUserInv_Account(ByVal UpdateAll As Boolean, _
                             ByVal UserIndex As Integer, _
                             ByVal Slot As Byte)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UpdateBanUserInv_Account_Err
        '</EhHeader>

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

100     With UserList(UserIndex).Account

            'Actualiza un solo slot
102         If Not UpdateAll Then

                'Actualiza el inventario
104             If .BancoInvent.Object(Slot).ObjIndex > 0 Then
106                 Call SendBanObj(UserIndex, Slot, .BancoInvent.Object(Slot), e_Account)
                Else
108                 Call SendBanObj(UserIndex, Slot, NullObj, e_Account)
                End If

            Else

                'Actualiza todos los slots
110             For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                    'Actualiza el inventario
112                 If .BancoInvent.Object(LoopC).ObjIndex > 0 Then
114                     Call SendBanObj(UserIndex, LoopC, .BancoInvent.Object(LoopC), e_Account)
                    Else
116                     Call SendBanObj(UserIndex, LoopC, NullObj, e_Account)
                    End If
            
118             Next LoopC

            End If

        End With

        '<EhFooter>
        Exit Sub

UpdateBanUserInv_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UpdateBanUserInv_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, _
                   ByVal i As Integer, _
                   ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserRetiraItem_Err
        '</EhHeader>


        Dim ObjIndex  As Integer

        Dim SlotEvent As Byte

100     If cantidad < 1 Then Exit Sub
    
102     Call WriteUpdateUserStats(UserIndex)
    
104     If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
        
106         If cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
            
108         ObjIndex = UserList(UserIndex).BancoInvent.Object(i).ObjIndex
        
110         SlotEvent = UserList(UserIndex).flags.SlotEvent
        
112         If SlotEvent > 0 Then
114             If Events(SlotEvent).LimitRed > 0 Then
116                 If ObjIndex = POCION_ROJA Then
118                     Call WriteConsoleMsg(UserIndex, "No puedes retirar pociones rojas en éste tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                End If
            
120             If Events(SlotEvent).ChangeClass > 0 Or Events(SlotEvent).ChangeRaze > 0 Or Events(SlotEvent).ChangeLevel > 0 Then
122                 If Events(SlotEvent).TimeCancel > 0 Then
124                     Call WriteConsoleMsg(UserIndex, "No puedes retirar objetos en este tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                End If
            End If
  
            'Agregamos el obj que compro al inventario
126         Call UserReciveObj(UserIndex, CInt(i), cantidad)
        
128         If ObjData(ObjIndex).Log = 1 Then
130             Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " retiró " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")

            End If

        End If
    
        'Actualizamos la ventana de comercio
132     Call UpdateVentanaBanco(UserIndex)

        '<EhFooter>
        Exit Sub

UserRetiraItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UserRetiraItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserRetiraItem_Account(ByVal UserIndex As Integer, _
                           ByVal i As Integer, _
                           ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserRetiraItem_Account_Err
        '</EhHeader>


        Dim ObjIndex  As Integer

        Dim SlotEvent As Byte
        
        
           If UserList(UserIndex).Account.Premium < 2 Then
                    Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior poseen un banco exclusivo. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If
                
100     If cantidad < 1 Then Exit Sub
    
102     Call WriteUpdateUserStats(UserIndex)
    
104     If UserList(UserIndex).Account.BancoInvent.Object(i).Amount > 0 Then
        
106         If cantidad > UserList(UserIndex).Account.BancoInvent.Object(i).Amount Then cantidad = UserList(UserIndex).Account.BancoInvent.Object(i).Amount
            
108         ObjIndex = UserList(UserIndex).Account.BancoInvent.Object(i).ObjIndex
        
110         SlotEvent = UserList(UserIndex).flags.SlotEvent
        
112         If SlotEvent > 0 Then
114             If Events(SlotEvent).LimitRed > 0 Then
116                 If ObjIndex = POCION_ROJA Then
118                     Call WriteConsoleMsg(UserIndex, "No puedes retirar pociones rojas en éste tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                End If
            
120             If Events(SlotEvent).ChangeClass > 0 Or Events(SlotEvent).ChangeRaze > 0 Or Events(SlotEvent).ChangeLevel > 0 Then
122                 If Events(SlotEvent).TimeCancel > 0 Then
124                     Call WriteConsoleMsg(UserIndex, "No puedes retirar objetos en este tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                End If
            End If
  
            'Agregamos el obj que compro al inventario
126         Call UserReciveObj_Account(UserIndex, CInt(i), cantidad)
        
128         If ObjData(ObjIndex).Log = 1 Then
130             Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " retiró " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")

            End If

        End If
    
        'Actualizamos la ventana de comercio
132     Call UpdateVentanaBanco(UserIndex)



        '<EhFooter>
        Exit Sub

UserRetiraItem_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UserRetiraItem_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, _
                  ByVal ObjIndex As Integer, _
                  ByVal cantidad As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserReciveObj_Err

        '</EhHeader>

        Dim Slot As Integer

        Dim obji As Integer

100     With UserList(UserIndex)

102         If .BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
    
104         obji = .BancoInvent.Object(ObjIndex).ObjIndex
    
            '¿Ya tiene un objeto de este tipo?
106         Slot = 1

108         Do Until .Invent.Object(Slot).ObjIndex = obji And .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
        
110             Slot = Slot + 1

112             If Slot > .CurrentInventorySlots Then

                    Exit Do

                End If

            Loop
    
            'Sino se fija por un slot vacio
114         If Slot > .CurrentInventorySlots Then
116             Slot = 1

118             Do Until .Invent.Object(Slot).ObjIndex = 0
120                 Slot = Slot + 1

122                 If Slot > .CurrentInventorySlots Then
124                     Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                Loop

126             .Invent.NroItems = .Invent.NroItems + 1

            End If
    
            'Mete el obj en el slot
128         If .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
                'Menor que MAX_INV_OBJS
130             .Invent.Object(Slot).ObjIndex = obji
132             .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + cantidad
        
134             Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), cantidad)
        
                'Actualizamos el inventario del usuario
136             Call UpdateUserInv(False, UserIndex, Slot)
        
                'Actualizamos el banco
138             Call UpdateBanUserInv(False, UserIndex, ObjIndex)
            Else
140             Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        '<EhFooter>
        Exit Sub

UserReciveObj_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserReciveObj " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub UserReciveObj_Account(ByVal UserIndex As Integer, _
                          ByVal ObjIndex As Integer, _
                          ByVal cantidad As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserReciveObj_Account_Err

        '</EhHeader>

        Dim Slot As Integer

        Dim obji As Integer
        
        If UserList(UserIndex).Account.Premium < 2 Then
            Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior poseen un banco exclusivo. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
                
100     With UserList(UserIndex)

102         If .Account.BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
    
104         obji = .Account.BancoInvent.Object(ObjIndex).ObjIndex
    
            '¿Ya tiene un objeto de este tipo?
106         Slot = 1

108         Do Until .Invent.Object(Slot).ObjIndex = obji And .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
        
110             Slot = Slot + 1

112             If Slot > .CurrentInventorySlots Then

                    Exit Do

                End If

            Loop
    
            'Sino se fija por un slot vacio
114         If Slot > .CurrentInventorySlots Then
116             Slot = 1

118             Do Until .Invent.Object(Slot).ObjIndex = 0
120                 Slot = Slot + 1

122                 If Slot > .CurrentInventorySlots Then
124                     Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                Loop

126             .Invent.NroItems = .Invent.NroItems + 1

            End If
    
            'Mete el obj en el slot
128         If .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
                'Menor que MAX_INV_OBJS
130             .Invent.Object(Slot).ObjIndex = obji
132             .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + cantidad
        
134             Call QuitarBancoInvItem_Account(UserIndex, CByte(ObjIndex), cantidad)
        
                'Actualizamos el inventario del usuario
136             Call UpdateUserInv(False, UserIndex, Slot)
        
                'Actualizamos el banco
138             Call UpdateBanUserInv_Account(False, UserIndex, ObjIndex)
            Else
140             Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        '<EhFooter>
        Exit Sub

UserReciveObj_Account_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserReciveObj_Account " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, _
                       ByVal Slot As Byte, _
                       ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuitarBancoInvItem_Err
        '</EhHeader>

        Dim ObjIndex As Integer

100     With UserList(UserIndex)
102         ObjIndex = .BancoInvent.Object(Slot).ObjIndex

            'Quita un Obj

104         .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount - cantidad
    
106         If .BancoInvent.Object(Slot).Amount <= 0 Then
108             .BancoInvent.NroItems = .BancoInvent.NroItems - 1
110             .BancoInvent.Object(Slot).ObjIndex = 0
112             .BancoInvent.Object(Slot).Amount = 0
            End If

        End With
    
        '<EhFooter>
        Exit Sub

QuitarBancoInvItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.QuitarBancoInvItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub QuitarBancoInvItem_Account(ByVal UserIndex As Integer, _
                               ByVal Slot As Byte, _
                               ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuitarBancoInvItem_Account_Err
        '</EhHeader>

        Dim ObjIndex As Integer

100     With UserList(UserIndex)
102         ObjIndex = .Account.BancoInvent.Object(Slot).ObjIndex

            'Quita un Obj

104         .Account.BancoInvent.Object(Slot).Amount = .Account.BancoInvent.Object(Slot).Amount - cantidad
    
106         If .Account.BancoInvent.Object(Slot).Amount <= 0 Then
108             .Account.BancoInvent.NroItems = .Account.BancoInvent.NroItems - 1
110             .Account.BancoInvent.Object(Slot).ObjIndex = 0
112             .Account.BancoInvent.Object(Slot).Amount = 0
            End If

        End With
    
        '<EhFooter>
        Exit Sub

QuitarBancoInvItem_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.QuitarBancoInvItem_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UpdateVentanaBanco(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UpdateVentanaBanco_Err
        '</EhHeader>

100     Call WriteBankOK(UserIndex)
        '<EhFooter>
        Exit Sub

UpdateVentanaBanco_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UpdateVentanaBanco " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, _
                     ByVal Item As Integer, _
                     ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserDepositaItem_Err
        '</EhHeader>


        Dim ObjIndex As Integer

100     If UserList(UserIndex).Invent.Object(Item).Amount > 0 And cantidad > 0 Then
    
102         If cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        
104         ObjIndex = UserList(UserIndex).Invent.Object(Item).ObjIndex
            
            If ObjData(ObjIndex).NoNada = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
106         If ObjData(ObjIndex).OBJType = otTransformVIP Then
108             If UserList(UserIndex).flags.TransformVIP = 1 Then
110                 Call TransformVIP_User(UserIndex, 0)
                End If
            End If
        
112         If ObjData(ObjIndex).OBJType = otGemaTelep Then
114             Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'Agregamos el obj que deposita al banco
116         Call UserDejaObj(UserIndex, CInt(Item), cantidad)
        
118         If ObjData(ObjIndex).Log = 1 Then
120             Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " depositó " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
            End If

        End If
    
        'Actualizamos la ventana del banco
122     Call UpdateVentanaBanco(UserIndex)

        '<EhFooter>
        Exit Sub

UserDepositaItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UserDepositaItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserDepositaItem_Account(ByVal UserIndex As Integer, _
                             ByVal Item As Integer, _
                             ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserDepositaItem_Account_Err
        '</EhHeader>

        Dim ObjIndex As Integer

100     If UserList(UserIndex).Invent.Object(Item).Amount > 0 And cantidad > 0 Then
    
102         If cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        
104         ObjIndex = UserList(UserIndex).Invent.Object(Item).ObjIndex
                
106         If ObjData(ObjIndex).OBJType = otTransformVIP Then
108             If UserList(UserIndex).flags.TransformVIP = 1 Then
110                 Call TransformVIP_User(UserIndex, 0)
                End If
            End If
        
112         If ObjData(ObjIndex).OBJType = otGemaTelep Then
114             Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            If ObjData(ObjIndex).NoNada = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'Agregamos el obj que deposita al banco
116         Call UserDejaObj_Account(UserIndex, CInt(Item), cantidad)
        
118         If ObjData(ObjIndex).Log = 1 Then
120             Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " depositó " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
            End If

        End If
    
        'Actualizamos la ventana del banco
122     Call UpdateVentanaBanco(UserIndex)

        '<EhFooter>
        Exit Sub

UserDepositaItem_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UserDepositaItem_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, _
                ByVal ObjIndex As Integer, _
                ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserDejaObj_Err
        '</EhHeader>

        Dim Slot As Integer

        Dim obji As Integer
    
100     If cantidad < 1 Then Exit Sub
    
102     With UserList(UserIndex)
104         obji = .Invent.Object(ObjIndex).ObjIndex
        
            '¿Ya tiene un objeto de este tipo?
106         Slot = 1

108         Do Until .BancoInvent.Object(Slot).ObjIndex = obji And .BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
110             Slot = Slot + 1
            
112             If Slot > MAX_BANCOINVENTORY_SLOTS Then

                    Exit Do

                End If

            Loop
        
            'Sino se fija por un slot vacio antes del slot devuelto
114         If Slot > MAX_BANCOINVENTORY_SLOTS Then
116             Slot = 1

118             Do Until .BancoInvent.Object(Slot).ObjIndex = 0
120                 Slot = Slot + 1
                
122                 If Slot > MAX_BANCOINVENTORY_SLOTS Then
124                     Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                Loop
            
126             .BancoInvent.NroItems = .BancoInvent.NroItems + 1
            End If
        
128         If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

                'Mete el obj en el slot
130             If .BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
                
                    'Menor que MAX_INV_OBJS
132                 .BancoInvent.Object(Slot).ObjIndex = obji
134                 .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount + cantidad
                
136                 Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), cantidad)
                        
                    'Actualizamos el inventario del usuario
138                 Call UpdateUserInv(False, UserIndex, ObjIndex)
                
                    'Actualizamos el inventario del banco
140                 Call UpdateBanUserInv(False, UserIndex, Slot)
                Else
142                 Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

UserDejaObj_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UserDejaObj " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub UserDejaObj_Account(ByVal UserIndex As Integer, _
                        ByVal ObjIndex As Integer, _
                        ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo UserDejaObj_Account_Err
        '</EhHeader>

        Dim Slot As Integer

        Dim obji As Integer
    
100     If cantidad < 1 Then Exit Sub
    
102     With UserList(UserIndex)
104         obji = .Invent.Object(ObjIndex).ObjIndex
        
            '¿Ya tiene un objeto de este tipo?
106         Slot = 1

108         Do Until .Account.BancoInvent.Object(Slot).ObjIndex = obji And .Account.BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
110             Slot = Slot + 1
            
112             If Slot > MAX_BANCOINVENTORY_SLOTS Then

                    Exit Do

                End If

            Loop
        
            'Sino se fija por un slot vacio antes del slot devuelto
114         If Slot > MAX_BANCOINVENTORY_SLOTS Then
116             Slot = 1

118             Do Until .Account.BancoInvent.Object(Slot).ObjIndex = 0
120                 Slot = Slot + 1
                
122                 If Slot > MAX_BANCOINVENTORY_SLOTS Then
124                     Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                Loop
            
126             .Account.BancoInvent.NroItems = .Account.BancoInvent.NroItems + 1
            End If
        
128         If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

                'Mete el obj en el slot
130             If .Account.BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
                
                    'Menor que MAX_INV_OBJS
132                 .Account.BancoInvent.Object(Slot).ObjIndex = obji
134                 .Account.BancoInvent.Object(Slot).Amount = .Account.BancoInvent.Object(Slot).Amount + cantidad
                
136                 Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), cantidad)
                        
                    'Actualizamos el inventario del usuario
138                 Call UpdateUserInv(False, UserIndex, ObjIndex)
                
                    'Actualizamos el inventario del banco
140                 Call UpdateBanUserInv_Account(False, UserIndex, Slot)
                Else
142                 Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

UserDejaObj_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.UserDejaObj_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserBovedaTxt_Err
        '</EhHeader>

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

104     For j = 1 To MAX_BANCOINVENTORY_SLOTS

106         If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If

        Next

        '<EhFooter>
        Exit Sub

SendUserBovedaTxt_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.SendUserBovedaTxt " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserBovedaTxt_Account(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserBovedaTxt_Account_Err
        '</EhHeader>

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).Account.BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

104     For j = 1 To MAX_BANCOINVENTORY_SLOTS

106         If UserList(UserIndex).Account.BancoInvent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).Account.BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Account.BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If

        Next

        '<EhFooter>
        Exit Sub

SendUserBovedaTxt_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.SendUserBovedaTxt_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserBovedaTxtFromChar_Err
        '</EhHeader>

        Dim j        As Integer

        Dim Charfile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long

100     Charfile = CharPath & charName & ".chr"

102     If FileExist(Charfile, vbNormal) Then
104         Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(Charfile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

108         For j = 1 To MAX_BANCOINVENTORY_SLOTS
110             Tmp = GetVar(Charfile, "BancoInventory", "Obj" & j)
112             ObjInd = ReadField(1, Tmp, Asc("-"))
114             ObjCant = ReadField(2, Tmp, Asc("-"))

116             If ObjInd > 0 Then
118                 Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
                End If

            Next

        Else
120         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

SendUserBovedaTxtFromChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.SendUserBovedaTxtFromChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub SendUserBovedaTxtFromChar_Account(ByVal sendIndex As Integer, _
                                      ByVal charName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo SendUserBovedaTxtFromChar_Account_Err
        '</EhHeader>

        Dim j        As Integer

        Dim Charfile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long
    
        Dim Account  As String
    
100     Charfile = CharPath & charName & ".chr"
    
102     Account = AccountPath & GetVar(Charfile, "INIT", "ACCOUNTNAME") & ACCOUNT_FORMAT
    
104     If FileExist(Account, vbNormal) Then
106         Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(Account, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

110         For j = 1 To MAX_BANCOINVENTORY_SLOTS
112             Tmp = GetVar(Account, "BancoInventory", "Obj" & j)
114             ObjInd = ReadField(1, Tmp, Asc("-"))
116             ObjCant = ReadField(2, Tmp, Asc("-"))

118             If ObjInd > 0 Then
120                 Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
                End If

            Next

        Else
122         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

SendUserBovedaTxtFromChar_Account_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modBanco.SendUserBovedaTxtFromChar_Account " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
