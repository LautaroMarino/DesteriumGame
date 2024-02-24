Attribute VB_Name = "Trabajo"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Const GASTO_ENERGIA_TRABAJADOR    As Byte = 2

Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 6

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo DoPermanecerOculto_Err

        '</EhHeader>

        '********************************************************
        'Autor: Nacho (Integer)
        'Last Modif: 11/19/2009
        'Chequea si ya debe mostrarse
        'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
        '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
        '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
        '********************************************************
        On Error GoTo ErrHandler

        Dim TiempoTranscurrido As Long
    
100     With UserList(UserIndex)
102         .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        
108         TiempoTranscurrido = (.Counters.TiempoOculto * frmMain.GameTimer.interval)
            
110         If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
116             Call WriteUpdateGlobalCounter(UserIndex, 1, .Counters.TiempoOculto / 40)

            End If

118
        
120         If .Counters.TiempoOculto <= 0 Then
122             If .Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                    ' Armaduras que permiten ocultarse por tiempo ilimitado
                    If .Invent.ArmourEqpObjIndex > 0 Then
126                     If ObjData(.Invent.ArmourEqpObjIndex).Oculto = 1 Then
128                         .Counters.TiempoOculto = IntervaloOculto
                            Exit Sub

                        End If
                
                    End If

                End If

130             .Counters.TiempoOculto = 0
132             .flags.Oculto = 0
            
134             If .flags.Navegando = 0 Then

136                 If .flags.Invisible = 0 Then
138                     Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    
                        'Si está en el oscuro no lo hacemos visible
140                     If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.zonaOscura Then
142                         Call SetInvisible(UserIndex, .Char.charindex, False)

                        End If

                    End If

                End If

            End If

        End With
    
        Exit Sub

ErrHandler:
144     Call LogError("Error en Sub DoPermanecerOculto")

        '<EhFooter>
        Exit Sub

DoPermanecerOculto_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoPermanecerOculto " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la fórmula y ahora anda bien.
    '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
    '***************************************************

    On Error GoTo ErrHandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(UserIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
                    
        If .Clase = eClass.Thief Then
            Suerte = 80

        End If
            
        res = RandomNumber(1, 100)
        
        If .Stats.MaxMan > 0 Then Suerte = Suerte / 2
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            
            If .Clase = eClass.Thief Then
                Suerte = Suerte * 2
            Else

                If .Stats.MaxMan > 0 Then
                    Suerte = Suerte / 2

                End If

            End If
            
            .Counters.TiempoOculto = Suerte
             
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.charindex, True)
                
                .PosOculto = .Pos
                Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                Call WriteUpdateGlobalCounter(UserIndex, 1, .Counters.TiempoOculto / 40)
                ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.Body = iFragataFantasmal


                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, Null)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
        Else
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, False)

        End If

        .Counters.Ocultando = .Counters.Ocultando + 1
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer, _
                    Optional ByVal NotRequired As Boolean = False)

        '***************************************************
        'Author: Unknown
        'Last Modification: 13/01/2010 (ZaMa)
        '13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
        '16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
        '10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la última barca que equipo era el galeón no explotaba(Y capaz no la tenía equipada :P).
        '***************************************************
        '<EhHeader>
        On Error GoTo DoNavega_Err

        '</EhHeader>

100     With UserList(UserIndex)

            If .Stats.Elv < 25 Then
                Call WriteConsoleMsg(UserIndex, "¡Las clases luchadoras pueden navegar a partir de Nivel 25!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
102         If NotRequired = False Then
104             If .Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
106                 Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
108                 Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    
                    Exit Sub
    
                End If
            
110             If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X - 1, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X + 1, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X, .Pos.Y - 1) = True And HayAgua(.Pos.Map, .Pos.X, .Pos.Y + 1) = True Then
112                 Call WriteConsoleMsg(UserIndex, "¡¡No puedes dejar de navegar en el agua!!", FontTypeNames.FONTTYPE_INFO)
    
                    Exit Sub
    
                End If

            End If
        
            ' No estaba navegando
114         If .flags.Navegando = 0 Then
116             .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
118             .Invent.BarcoSlot = Slot
            
120             .Char.Head = 0
            
                ' No esta muerto
122             If .flags.Muerto = 0 Then
            
124                 Call ToggleBoatBody(UserIndex)
                
                    ' Pierde el ocultar
126                 If .flags.Oculto = 1 Then
128                     .flags.Oculto = 0
130                     Call SetInvisible(UserIndex, .Char.charindex, False)
132                     Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                    End If
               
                    ' Siempre se ve la barca (Nunca esta invisible), pero solo para el cliente.
134                 If .flags.Invisible = 1 Then
136                     Call SetInvisible(UserIndex, .Char.charindex, False)
138                     UserList(UserIndex).Counters.DrawersCount = 0

                    End If
                
                    ' Esta muerto
                Else
140                 .Char.Body = iFragataFantasmal
142                 .Char.ShieldAnim = NingunEscudo
144                 .Char.WeaponAnim = NingunArma
146                 .Char.CascoAnim = NingunCasco

                End If
            
                ' Comienza a navegar
148             .flags.Navegando = 1
        
                ' Estaba navegando
            Else
150             .Invent.BarcoObjIndex = 0
152             .Invent.BarcoSlot = 0
        
                ' No esta muerto
154             If .flags.Muerto = 0 Then
156                 .Char.Head = .OrigChar.Head

158                 If .Invent.ArmourEqpObjIndex > 0 Then
160                     .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
                    Else
162                     Call DarCuerpoDesnudo(UserIndex)

                    End If
                
164                 If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = GetShieldAnim(UserIndex, .Invent.EscudoEqpObjIndex)

166                 If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, .Invent.WeaponEqpObjIndex)

168                 If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = GetHelmAnim(UserIndex, .Invent.CascoEqpObjIndex)
                
                    ' Al dejar de navegar, si estaba invisible actualizo los clientes
170                 If .flags.Invisible = 1 Then
172                     Call SetInvisible(UserIndex, .Char.charindex, True)
174                     UserList(UserIndex).Counters.DrawersCount = RandomNumberPower(1, 200)

                    End If
                
                    ' Esta muerto
                Else
176                 .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
180                 .Char.Head = iCabezaMuerto(Escriminal(UserIndex))

182                 .Char.ShieldAnim = NingunEscudo
184                 .Char.WeaponAnim = NingunArma
186                 .Char.CascoAnim = NingunCasco

                End If
            
                ' Termina de navegar
188             .flags.Navegando = 0

            End If
        
            ' Actualizo clientes
190         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

        End With
    
192     Call WriteNavigateToggle(UserIndex)

        '<EhFooter>
        Exit Sub

DoNavega_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoNavega " & "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) Then
                Call DoLingotes(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de minería suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en FundirMineral. Error " & Err.number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Long, _
                      ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 10/07/2010
        '10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
        '***************************************************
        '<EhHeader>
        On Error GoTo TieneObjetos_Err
        '</EhHeader>

        Dim i     As Integer

        Dim Total As Long

100     For i = 1 To UserList(UserIndex).CurrentInventorySlots

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
104             Total = Total + UserList(UserIndex).Invent.Object(i).Amount
            End If

106     Next i
    
108     If cant <= Total Then
110         TieneObjetos = True

            Exit Function

        End If
        
        '<EhFooter>
        Exit Function

TieneObjetos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.TieneObjetos " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

'# Los objetos [BRONCE] [PLATA] [ORO] [PREMIUM] Se consideran especiales.
Function TieneObjetos_Especiales(ByVal UserIndex As Integer, _
                                 ByVal Bronce As Byte, _
                                 ByVal Plata As Byte, _
                                 ByVal Oro As Byte, _
                                 ByVal Premium As Byte) As String
        '<EhHeader>
        On Error GoTo TieneObjetos_Especiales_Err
        '</EhHeader>

        Dim A     As Integer
        Dim ObjIndex As Integer
    
        Dim Total As Long

100     For A = 1 To UserList(UserIndex).CurrentInventorySlots
102         ObjIndex = UserList(UserIndex).Invent.Object(A).ObjIndex
        
104         If ObjIndex > 0 Then
106             If Bronce = 0 And ObjData(ObjIndex).Bronce = 1 Then
108                 TieneObjetos_Especiales = "El evento no permite los objetos [AVENTURERO]"
                    Exit Function
                End If
            
110             If Plata = 0 And ObjData(ObjIndex).Plata = 1 Then
112                 TieneObjetos_Especiales = "El evento no permite los objetos [HEROE]"
                    Exit Function
                End If
            
114             If Oro = 0 And ObjData(ObjIndex).Oro = 1 Then
116                 TieneObjetos_Especiales = "El evento no permite los objetos [LEYENDA]"
                    Exit Function
                End If
            
118             If Premium = 0 And ObjData(ObjIndex).Premium = 1 Then
120                 TieneObjetos_Especiales = "El evento no permite los objetos [PREMIUM]"
                    Exit Function
                End If
            End If
122     Next A
        
        '<EhFooter>
        Exit Function

TieneObjetos_Especiales_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.TieneObjetos_Especiales " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Long, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************
        '<EhHeader>
        On Error GoTo QuitarObjetos_Err
        '</EhHeader>

        Dim i As Integer
100     For i = 1 To UserList(UserIndex).CurrentInventorySlots
102         With UserList(UserIndex).Invent.Object(i)
104             If .ObjIndex = ItemIndex Then
106                 If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                
108                 .Amount = .Amount - cant
                
110                 If .Amount <= 0 Then
112                     cant = Abs(.Amount)
114                     UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
116                     .Amount = 0
118                     .ObjIndex = 0
                    Else
120                     cant = 0
                    End If
                
122                 Call UpdateUserInv(False, UserIndex, i)
                
124                 If cant = 0 Then Exit Sub
                End If
            End With
126     Next i

        '<EhFooter>
        Exit Sub

QuitarObjetos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.QuitarObjetos " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub QuitarObjetoEspecifico(ByVal ItemIndex As Integer, ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo QuitarObjetoEspecifico_Err
        '</EhHeader>
        Dim i As Integer
    
100     For i = 1 To UserList(UserIndex).CurrentInventorySlots
102         With UserList(UserIndex).Invent.Object(i)
104             If .ObjIndex = ItemIndex Then
106                 If .Equipped = 1 Then Call Desequipar(UserIndex, i)

108                 UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
110                 .Amount = 0
112                 .ObjIndex = 0
                
114                 Call UpdateUserInv(False, UserIndex, i)
                
                End If
            End With
116     Next i

        '<EhFooter>
        Exit Sub

QuitarObjetoEspecifico_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.QuitarObjetoEspecifico " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
        '<EhHeader>
        On Error GoTo MineralesParaLingote_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
100     Select Case Lingote

            Case iMinerales.HierroCrudo
102             MineralesParaLingote = 13

104         Case iMinerales.PlataCruda
106             MineralesParaLingote = 25

108         Case iMinerales.OroCrudo
110             MineralesParaLingote = 50

112         Case Else
114             MineralesParaLingote = 10000
        End Select

        '<EhFooter>
        Exit Function

MineralesParaLingote_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.MineralesParaLingote " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo DoLingotes_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 16/11/2009
        '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
        '***************************************************
        '    Call LogTarea("Sub DoLingotes")
        Dim Slot           As Integer

        Dim obji           As Integer

        Dim CantidadItems  As Integer

        Dim TieneMinerales As Boolean

        Dim OtroUserIndex  As Integer
    
100     With UserList(UserIndex)

102         If .flags.Comerciando Then
104             OtroUserIndex = .ComUsu.DestUsu
                
106             If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
108                 Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
110                 Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
112                 Call LimpiarComercioSeguro(UserIndex)
114                 Call Protocol.FlushBuffer(OtroUserIndex)
                End If
            End If
        
116         CantidadItems = MaximoInt(1, CInt((.Stats.Elv - 4) / 5))

118         Slot = .flags.TargetObjInvSlot
120         obji = .Invent.Object(Slot).ObjIndex
        
122         While CantidadItems > 0 And Not TieneMinerales

124             If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
126                 TieneMinerales = True
                Else
128                 CantidadItems = CantidadItems - 1
                End If

            Wend
        
130         If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
132             Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
134         .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems

136         If .Invent.Object(Slot).Amount < 1 Then
138             .Invent.Object(Slot).Amount = 0
140             .Invent.Object(Slot).ObjIndex = 0
            End If
        
            Dim MiObj As Obj

142         MiObj.Amount = CantidadItems
144         MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

146         If Not MeterItemEnInventario(UserIndex, MiObj) Then
148             Call TirarItemAlPiso(.Pos, MiObj)
            End If
        
150         Call UpdateUserInv(False, UserIndex, Slot)
152         Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & CantidadItems & " lingote" & IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
154         .Counters.Trabajando = .Counters.Trabajando + 1
        End With

        '<EhFooter>
        Exit Sub

DoLingotes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoLingotes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoFundir(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo DoFundir_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 03/06/2010
        '03/06/2010 - Pato: Si es el último ítem a fundir y está equipado lo desequipamos.
        '11/03/2010 - ZaMa: Reemplazo división por producto para uan mejor performanse.
        '***************************************************
        Dim i             As Integer

        Dim Num           As Integer

        Dim Slot          As Byte

        Dim Lingotes(2)   As Integer

        Dim OtroUserIndex As Integer

100     With UserList(UserIndex)

102         If .flags.Comerciando Then
104             OtroUserIndex = .ComUsu.DestUsu
                
106             If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
108                 Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
110                 Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
112                 Call LimpiarComercioSeguro(UserIndex)
114                 Call Protocol.FlushBuffer(OtroUserIndex)
                End If
            End If
        
116         Slot = .flags.TargetObjInvSlot
        
118         With .Invent.Object(Slot)
120             .Amount = .Amount - 1
            
122             If .Amount < 1 Then
124                 If .Equipped = 1 Then Call Desequipar(UserIndex, Slot)
                
126                 .Amount = 0
128                 .ObjIndex = 0
                End If

            End With
        
130         Num = RandomNumber(10, 25)
        
132         Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01
134         Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01
136         Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01
    
            Dim MiObj(2) As Obj
        
138         For i = 0 To 2
140             MiObj(i).Amount = Lingotes(i)
142             MiObj(i).ObjIndex = LingoteHierro + i 'Una gran negrada pero práctica
            
144             If MiObj(i).Amount > 0 Then
146                 If Not MeterItemEnInventario(UserIndex, MiObj(i)) Then
148                     Call TirarItemAlPiso(.Pos, MiObj(i))
                    End If
                End If

150         Next i
        
152         Call UpdateUserInv(False, UserIndex, Slot)
154         Call WriteConsoleMsg(UserIndex, "¡Has obtenido el " & Num & "% de los lingotes utilizados para la construcción del objeto!", FontTypeNames.FONTTYPE_INFO)
    
156         .Counters.Trabajando = .Counters.Trabajando + 1
        End With

        '<EhFooter>
        Exit Sub

DoFundir_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoFundir " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


Sub DoAdminInvisible(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2010 (ZaMa)
        'Makes an admin invisible o visible.
        '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
        '***************************************************
        '<EhHeader>
        On Error GoTo DoAdminInvisible_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

102         If .flags.AdminInvisible = 0 Then

                ' Sacamos el mimetizmo
104             If .flags.Mimetizado = 1 Then
106                 .Char.Body = .CharMimetizado.Body
108                 .Char.Head = .CharMimetizado.Head
110                 .Char.CascoAnim = .CharMimetizado.CascoAnim
112                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
114                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
116                 .Counters.Mimetismo = 0
118                 .flags.Mimetizado = 0
                    ' Se fue el efecto del mimetismo, puede ser atacado por npcs
120                 .flags.Ignorado = False
                End If
            
122             .flags.AdminInvisible = 1
124             .flags.Invisible = 1
126             .flags.Oculto = 1
128             .flags.OldBody = .Char.Body
130             .flags.OldHead = .Char.Head
132             .Char.Body = 0
134             .Char.Head = 0
            
                
                ' Solo el admin sabe que se hace invi
136             Call SendData(ToOne, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
                'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
138             Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.charindex))
               ' Call ModAreas.DeleteEntity(UserIndex, ENTITY_TYPE_PLAYER)
        
            Else
140             .flags.AdminInvisible = 0
142             .flags.Invisible = 0
144             .flags.Oculto = 0
146             .Counters.TiempoOculto = 0
148             .Char.Body = .flags.OldBody
150             .Char.Head = .flags.OldHead
            
                ' Solo el admin sabe que se hace visible
152             Call SendData(ToOne, UserIndex, PrepareMessageCharacterChange(.Char.Body, 0, .Char.Head, .Char.Heading, .Char.charindex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim, .Char.AuraIndex, .flags.ModoStream, False, False))
154             Call SendData(ToOne, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
            
                'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
156             Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, True, True)
            
                 ' Se lo mando a los demas
158             Call ModAreas.CreateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, ModAreas.DEFAULT_ENTITY_WIDTH, ModAreas.DEFAULT_ENTITY_HEIGHT)
        
            End If

        End With
    
        '<EhFooter>
        Exit Sub

DoAdminInvisible_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoAdminInvisible " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub DoMineria(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 28/05/2010
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim CantidadItems As Integer

    With UserList(UserIndex)

        ' Si estaba oculto, se vuelve visible
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
                
            If .flags.Invisible = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Call SetInvisible(UserIndex, .Char.charindex, False)
            End If
        End If
            
        Call QuitarSta(UserIndex, RandomNumber(0, EsfuerzoTalarLeñador))
    
        Dim Skill As Integer

        Skill = .Stats.UserSkills(eSkill.Mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= 5 Then

            Dim MiObj As Obj
        
            If .flags.TargetObj = 0 Then Exit Sub
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
            CantidadItems = MaxItemsExtraibles(.Stats.Elv)
            
            MiObj.Amount = RandomNumber(1, CantidadItems + ObjData(.Invent.WeaponEqpObjIndex).ProbPesca)

        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
        
            Call SubirSkill(UserIndex, eSkill.Mineria, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9
            End If

            '[/CDT]
            Call SubirSkill(UserIndex, eSkill.Mineria, False)
        End If
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_MINERO, .Pos.X, .Pos.Y))
        
        If Not Escriminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoMineria")

End Sub
Public Sub DoPescar(ByVal UserIndex As Integer, ByVal WeaponIndex As Integer)
        '<EhHeader>
        On Error GoTo DoPescar_Err
        '</EhHeader>

        Dim iSkill        As Integer

        Dim Suerte        As Integer

        Dim res           As Integer

        Dim CantidadItems As Integer
    
        Dim LastFish As Byte    ' Ultimo pescado disponible
        Dim MaxSuerte As Byte   ' Mejora la cantidad de suerte segun el barco que tenga
     
100     With UserList(UserIndex)
            
            If MapInfo(.Pos.Map).Pesca = 0 Then Exit Sub ' # No hay peces en el mapa
            
104         Select Case WeaponIndex
                Case CAÑA_PESCA
                    If .Invent.BarcoObjIndex > 0 Then
                        MaxSuerte = ObjData(.Invent.BarcoObjIndex).ProbPesca
                    End If
                      
110             Case RED_PESCA
112                 If .Invent.BarcoObjIndex <> 475 And .Invent.BarcoObjIndex <> 476 Then Exit Sub

116                 If Abs(.Pos.X - .flags.TargetX) + Abs(.Pos.Y - .flags.TargetY) > 5 Then
118                     Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                            
120                 If .Pos.X = .flags.TargetX And .Pos.Y = .flags.TargetY Then
122                     Call WriteConsoleMsg(UserIndex, "No puedes pescar desde allí.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                
126                 MaxSuerte = ObjData(.Invent.BarcoObjIndex).ProbPesca
128             Case Else
                    Exit Sub
            End Select
                            
            'Play sound!
130         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PESCAR, .Pos.X, .Pos.Y, .Char.charindex))
                              
132         Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
        
134         iSkill = .Stats.UserSkills(eSkill.Pesca)
        
136         Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 55)

138         If Suerte > 0 Then
140             res = RandomNumber(1, Suerte)
            
142             If res <= (6 + MaxSuerte) Then
            
                    Dim MiObj                        As Obj

                    Dim A                            As Long

                    Dim N                            As Long
                
                    Dim SacoPez                      As Boolean
                    
                    Dim Slot                         As Integer
                    
                    Slot = RandomNumber(1, MapInfo(.Pos.Map).Pesca)

146                 MiObj.ObjIndex = MapInfo(.Pos.Map).PescaItem(Slot)
148                 MiObj.Amount = RandomNumber(1, ObjData(WeaponIndex).ProbPesca)
                        
150                 If RandomNumber(1, 100) <= ObjData(MiObj.ObjIndex).ProbPesca Then

152                     If Not MeterItemEnInventario(UserIndex, MiObj) Then
154                         Call TirarItemAlPiso(.Pos, MiObj)
                        End If
                            
                        Call SubirSkill(UserIndex, eSkill.Pesca, True)

                    End If
158
                Else
                
164                 Call SubirSkill(UserIndex, eSkill.Pesca, False)
                End If
            End If
        
166         .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

168         If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
170         .Counters.Trabajando = .Counters.Trabajando + 1
        End With
    

        '<EhFooter>
        Exit Sub

DoPescar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoPescar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 05/04/2010
        'Last Modification By: ZaMa
        '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
        '27/11/2009: ZaMa - Optimizacion de codigo.
        '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
        '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
        '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
        '23/04/2010: ZaMa - No se puede robar mas sin energia.
        '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
        '*************************************************
        '<EhHeader>
        On Error GoTo DoRobar_Err
        '</EhHeader>

        Dim OtroUserIndex As Integer

100     If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
        ' Caos robando a caos?
102     If UserList(LadrOnIndex).flags.Oculto = 0 Then
104         Call WriteConsoleMsg(LadrOnIndex, "¡No puedes robar o hurtar objetos si no te encuentras oculto!", FontTypeNames.FONTTYPE_FIGHT)

            Exit Sub

        End If
        
106     If UserList(VictimaIndex).flags.EnConsulta Then
108         Call WriteConsoleMsg(LadrOnIndex, "¡¡¡No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
110     With UserList(LadrOnIndex)
    
112         If .flags.Seguro Then
114             If Not Escriminal(VictimaIndex) Then
116                 Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)

                    Exit Sub

                End If

            Else

118             If .Faction.Status = r_Armada Then
120                 If Not Escriminal(VictimaIndex) Then
122                     Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejército real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)

                        Exit Sub

                    End If

                End If

            End If
        
            ' Caos robando a caos?
124         If UserList(VictimaIndex).Faction.Status = r_Caos And .Faction.Status = r_Caos Then
126             Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legión oscura.", FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
        
128         If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
            ' Tiene energia?
130         If .Stats.MinSta < 15 Then
132             If .Genero = eGenero.Hombre Then
134                 Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
                Else
136                 Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
                Exit Sub

            End If
            
            ' ¿La victima tiene energia para ser robado?
            If UserList(VictimaIndex).Stats.MinSta < 15 Then
                Call WriteConsoleMsg(LadrOnIndex, "El Personaje está muy cansado para poder defenderse del hurto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
        
            ' Quito energia
138         Call QuitarSta(LadrOnIndex, 15)
        
            Dim GuantesHurto As Boolean
    
140         If .Invent.WeaponEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
142         If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
                Dim Suerte     As Integer

                Dim res        As Integer

                Dim RobarSkill As Byte
            
144             RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
146             If RobarSkill <= 10 Then
148                 Suerte = 35
150             ElseIf RobarSkill <= 20 Then
152                 Suerte = 30
154             ElseIf RobarSkill <= 30 Then
156                 Suerte = 28
158             ElseIf RobarSkill <= 40 Then
160                 Suerte = 24
162             ElseIf RobarSkill <= 50 Then
164                 Suerte = 22
166             ElseIf RobarSkill <= 60 Then
168                 Suerte = 20
170             ElseIf RobarSkill <= 70 Then
172                 Suerte = 18
174             ElseIf RobarSkill <= 80 Then
176                 Suerte = 15
178             ElseIf RobarSkill <= 90 Then
180                 Suerte = 10
182             ElseIf RobarSkill < 100 Then
184                 Suerte = 7
                Else
186                 Suerte = 5

                End If
            
188             res = RandomNumber(1, Suerte)
                
190             If res < 3 Then 'Exito robo
192                 If UserList(VictimaIndex).flags.Comerciando Then
194                     OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
196                     If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
198                         Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
200                         Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
202                         Call LimpiarComercioSeguro(VictimaIndex)
204                         Call Protocol.FlushBuffer(OtroUserIndex)

                        End If

                    End If
               
206                 If (RandomNumber(1, 100) < 35) Then
208                     If TieneObjetosRobables(VictimaIndex) Then
210                         Call RobarObjeto(LadrOnIndex, VictimaIndex)
                        Else
212                         Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else 'Roba oro

214                     If UserList(VictimaIndex).Stats.Gld > 0 Then

                            Dim N As Long
                        
216                         If .Clase = eClass.Thief Then

                                ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
218                             If GuantesHurto Then
220                                 N = RandomNumber((.Stats.Elv) * 50, (.Stats.Elv) * 100)
                                Else
222                                 N = RandomNumber(.Stats.Elv * 25, (.Stats.Elv) * 50)

                                End If
                            
224                             If UserList(VictimaIndex).flags.Paralizado = 1 Or UserList(VictimaIndex).flags.Inmovilizado = 1 Then
226                                 N = N * 1.3

                                End If
                            
                            Else
228                             N = RandomNumber(1, 100)

                            End If

230                         If N > UserList(VictimaIndex).Stats.Gld Then
232                             N = UserList(VictimaIndex).Stats.Gld
234                             Call WriteConsoleMsg(LadrOnIndex, "¡Le has robado todo el Oro a " & UserList(VictimaIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
236                             Call WriteConsoleMsg(VictimaIndex, "¡Te han robado todo el Oro!", FontTypeNames.FONTTYPE_INFO)
                            Else
238                             Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
240                             Call WriteConsoleMsg(VictimaIndex, "Te han robado " & N & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)

                            End If
                        
242                         UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - N
                        
244                         .Stats.Gld = .Stats.Gld + N

246                         If .Stats.Gld > MAXORO Then .Stats.Gld = MAXORO
                      
248                         Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
250                         Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
252                         Call FlushBuffer(VictimaIndex)
                        Else
254                         Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                
256                 Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
                Else
258                 Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
260                 Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
262                 Call FlushBuffer(VictimaIndex)
                
264                 Call SubirSkill(LadrOnIndex, eSkill.Robar, False)

                End If
        
266             If Not Escriminal(LadrOnIndex) Then
268                 If Not Escriminal(VictimaIndex) Then
270                     Call VolverCriminal(LadrOnIndex)

                    End If

                End If
            
                ' Se pudo haber convertido si robo a un ciuda
272             If Escriminal(LadrOnIndex) Then
274                 .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron

276                 If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP

                End If

            End If

        End With


        '<EhFooter>
        Exit Sub

DoRobar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoRobar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal Slot As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        ' Agregué los barcos
        ' Esta funcion determina qué objetos son robables.
        ' 22/05/2010: Los items newbies ya no son robables.
        '***************************************************
        '<EhHeader>
        On Error GoTo ObjEsRobable_Err
        '</EhHeader>

        Dim OI As Integer

100     OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

102     ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And _
           UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
           ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And _
           ObjData(OI).OBJType <> eOBJType.otBarcos And _
           ObjData(OI).OBJType <> eOBJType.otMonturas And _
           ObjData(OI).Bronce <> 1 And _
           ObjData(OI).Premium <> 1 And _
           ObjData(OI).Plata <> 1 And _
           ObjData(OI).Oro <> 1 And Not ItemNewbie(OI) And _
           ObjData(OI).NoNada <> 1 And _
           Not ObjData(OI).OBJType = otGemaTelep _
           And Not ObjData(OI).OBJType = otTransformVIP

        '<EhFooter>
        Exit Function

ObjEsRobable_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.ObjEsRobable " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
        '***************************************************
        '<EhHeader>
        On Error GoTo RobarObjeto_Err
        '</EhHeader>

        Dim flag As Boolean

        Dim i    As Integer

100     flag = False

102     With UserList(VictimaIndex)

104         If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
106             i = 1

108             Do While Not flag And i <= .CurrentInventorySlots

                    'Hay objeto en este slot?
110                 If .Invent.Object(i).ObjIndex > 0 Then
112                     If ObjEsRobable(VictimaIndex, i) Then
114                         If RandomNumber(1, 10) < 4 Then flag = True

                        End If

                    End If

116                 If Not flag Then i = i + 1
                Loop

            Else
118             i = .CurrentInventorySlots

120             Do While Not flag And i > 0

                    'Hay objeto en este slot?
122                 If .Invent.Object(i).ObjIndex > 0 Then
124                     If ObjEsRobable(VictimaIndex, i) Then
126                         If RandomNumber(1, 10) < 4 Then flag = True

                        End If

                    End If

128                 If Not flag Then i = i - 1
                Loop

            End If
    
130         If flag Then

                Dim MiObj     As Obj

                Dim Num       As Integer

                Dim ObjAmount As Integer
        
132             ObjAmount = .Invent.Object(i).Amount

134             If UserList(VictimaIndex).flags.Paralizado = 1 Or UserList(VictimaIndex).flags.Inmovilizado = 1 Then
                     'Cantidad al azar entre el 15% y el 20% del total, con minimo 1.
136                 Num = MaximoInt(1, RandomNumber(ObjAmount * 0.15, ObjAmount * 0.2))
                Else
                    'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
138                 Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))

                End If

140             MiObj.Amount = Num
142             MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
144             .Invent.Object(i).Amount = ObjAmount - Num
                    
146             If .Invent.Object(i).Amount <= 0 Then
148                 Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

                End If
                
150             Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
152             If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
154                 Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

                End If
        
156             If UserList(LadrOnIndex).Clase = eClass.Thief Then
158                 Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
160                 Call WriteConsoleMsg(VictimaIndex, "Te han robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
                Else
162                 Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)

                End If

            Else
164             Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)

            End If

            'If exiting, cancel de quien es robado
166         Call CancelExit(VictimaIndex)

        End With

        '<EhFooter>
        Exit Sub

RobarObjeto_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.RobarObjeto " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long, ByRef exito As Boolean)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el daño
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Suerte As Integer
    Dim Skill  As Integer
    Dim ObjIndex As Integer
    
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

    Select Case UserList(UserIndex).Clase

        Case eClass.Assasin
            Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.078) * Skill + 4.45)
          
        Case eClass.Cleric, eClass.Paladin
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
          
        Case eClass.Bard, eClass.Hunter, eClass.Warrior
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
          
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)
    End Select


    ' Si no es Daga Arrojadiza
    
    Dim WeaponIndex As Integer
    Dim CanDaga As Boolean
    
    WeaponIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
    CanDaga = True
    
    If WeaponIndex > 0 Then
        If ObjData(WeaponIndex).proyectil = 1 And ObjData(WeaponIndex).Apuñala = 1 Then
            CanDaga = False
        End If
    End If
    
    If CanDaga Then
        If VictimUserIndex Then
            If UserList(UserIndex).Clase = eClass.Assasin Then
                If UserList(UserIndex).Char.Heading = UserList(VictimUserIndex).Char.Heading Then Suerte = 60
            End If
        End If
    End If
    
    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            daño = Round(daño * 1.5, 0)

            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
            
            
            SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, UserList(UserIndex).DañoApu + daño, eDamageType.d_Apuñalar)
            Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessagePlayEffect(eSound.sApuñaladaEspalda, UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y))
            
            Call WriteConsoleMsg(VictimUserIndex, "¡Te han dado una apuñalada por " & Int(UserList(UserIndex).DañoApu + daño) & "!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(UserIndex, "¡Apuñalada por " & Int(UserList(UserIndex).DañoApu + daño) & "!", FontTypeNames.FONTTYPE_FIGHT)
            
            Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateFX(UserList(VictimUserIndex).Char.charindex, FXIDs.FX_APUÑALADA, 1))
            
            exito = True
        Else
            ObjIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
            
            daño = daño + ObjData(ObjIndex).NpcBonusDamage
            
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            SendData SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateDamage(Npclist(VictimNpcIndex).Pos.X, Npclist(VictimNpcIndex).Pos.Y, Int(UserList(UserIndex).DañoApu + daño), eDamageType.d_Apuñalar)
            Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessagePlayEffect(eSound.sApuñaladaEspalda, Npclist(VictimNpcIndex).Pos.X, Npclist(VictimNpcIndex).Pos.Y))
            Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(UserList(UserIndex).DañoApu + daño), FontTypeNames.FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateFX(Npclist(VictimNpcIndex).Char.charindex, FXIDs.FX_APUÑALADA, 1))

            '[Alejo]
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
            
            exito = True
        End If
          
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
    Else
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
      '  Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
        'SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, daño, DAMAGE_NORMAL)
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
    End If

    Exit Sub
ErrHandler:
End Sub

Public Sub DoAcuchillar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 12/01/2010
        '***************************************************
        '<EhHeader>
        On Error GoTo DoAcuchillar_Err
        '</EhHeader>

100     If RandomNumber(1, 100) <= PROB_ACUCHILLAR Then
102         daño = Int(daño * DAÑO_ACUCHILLAR)
        
104         If VictimUserIndex <> 0 Then
        
106             With UserList(VictimUserIndex)
108                 .Stats.MinHp = .Stats.MinHp - daño
110                 Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & .Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
112                 Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha acuchillado por " & daño, FontTypeNames.FONTTYPE_FIGHT)
                End With
            
            Else
        
114             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
116             Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
118             Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
        
            End If
        End If
    
        '<EhFooter>
        Exit Sub

DoAcuchillar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoAcuchillar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoGolpeCritico_Npcs(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal daño As Long)
        '***************************************************
        'Autor: Lautaro
        'Last Modification:
        ' Lo hacemos aparte porque queremos dejar el otro
        '***************************************************
        '<EhHeader>
        On Error GoTo DoGolpeCritico_Npcs_Err
        '</EhHeader>
   
100     daño = UserList(UserIndex).Stats.Elv * 2
        
102     With UserList(UserIndex)
        
104         If .Clase <> eClass.Warrior And .Clase <> eClass.Hunter And .Clase <> eClass.Thief Then Exit Sub
    
        End With
    
106     With Npclist(NpcIndex)
    
108         .Stats.MinHp = .Stats.MinHp - daño
        
           ' Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_CRITICO)
110         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y - 1, daño, eDamageType.d_DañoNpc_Critical))
112         Call CalcularDarExp(UserIndex, NpcIndex, daño)
        End With
    
        '<EhFooter>
        Exit Sub

DoGolpeCritico_Npcs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoGolpeCritico_Npcs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long)
        '<EhHeader>
        On Error GoTo DoGolpeCritico_Err
        '</EhHeader>

        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 28/01/2007
        '01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
        '***************************************************
        Dim Suerte      As Integer

        Dim Skill       As Integer

        Dim WeaponIndex As Integer
    
        Exit Sub
    
100     With UserList(UserIndex)
            ' Es bandido?
            'If .Clase <> eClass.Bandit Then Exit Sub
        
102         WeaponIndex = .Invent.WeaponEqpObjIndex
        
            ' Es una espada vikinga?
104         If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
106         Skill = .Stats.UserSkills(eSkill.Armas)
        End With
    
108     Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
110     If RandomNumber(1, 100) <= Suerte Then
    
112         daño = Int(daño * 0.75)
        
114         If VictimUserIndex <> 0 Then
            
116             With UserList(VictimUserIndex)
118                 .Stats.MinHp = .Stats.MinHp - daño
120                 Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & .Name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
122                 Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
                End With
            
            Else
        
124             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
126             Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
128             Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
            
            End If
        
        End If

        '<EhFooter>
        Exit Sub

DoGolpeCritico_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoGolpeCritico " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuitarSta_Err
        '</EhHeader>

        On Error GoTo ErrHandler

100     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - cantidad

102     If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
104     Call WriteUpdateSta(UserIndex)
    
        Exit Sub

ErrHandler:
106     Call LogError("Error en QuitarSta. Error " & Err.number & " : " & Err.description)
    
        '<EhFooter>
        Exit Sub

QuitarSta_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.QuitarSta " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, _
                   ByVal ObjIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 28/05/2010
    '16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim CantidadItems As Integer

    Dim Skill         As Integer

    With UserList(UserIndex)

        Call QuitarSta(UserIndex, RandomNumber(0, EsfuerzoTalarLeñador))
    
        Skill = .Stats.UserSkills(eSkill.Talar)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= 4 Then

            Dim MiObj As Obj
            CantidadItems = MaxItemsExtraibles(.Stats.Elv)
            
            MiObj.Amount = RandomNumber(1, CantidadItems)
            MiObj.ObjIndex = ObjIndex
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)
            End If
        
            Call SubirSkill(UserIndex, eSkill.Talar, True)
        Else

            '[/CDT]
            Call SubirSkill(UserIndex, eSkill.Talar, False)
        End If
    
        If Not Escriminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo DoMeditar_Err
        '</EhHeader>


        Dim Mana As Long
        
100     With UserList(UserIndex)

102         .Counters.TimerMeditar = .Counters.TimerMeditar + 1
104         .Counters.TiempoInicioMeditar = .Counters.TiempoInicioMeditar + 1
            
106         If .Counters.TimerMeditar >= IntervaloMeditar Then

108             Mana = Porcentaje(.Stats.MaxMan, Porcentaje(Balance.PorcentajeRecuperoMana, 50 + .Stats.UserSkills(eSkill.Magia) * 0.5))

110             If Mana <= 0 Then Mana = 1

112             If .Stats.MinMan + Mana >= .Stats.MaxMan Then

114                 .Stats.MinMan = .Stats.MaxMan
116                 .flags.Meditando = False
118                 .Char.FX = 0

120                 Call WriteUpdateMana(UserIndex)

124                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
                
                Else
                    
126                 .Stats.MinMan = .Stats.MinMan + Mana
128                 Call WriteUpdateMana(UserIndex)

                End If

132             .Counters.TimerMeditar = 0

            End If

        End With
        


        '<EhFooter>
        Exit Sub

DoMeditar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoMeditar " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modif: 15/04/2010
        'Unequips either shield, weapon or helmet from target user.
        '***************************************************
        '<EhHeader>
        On Error GoTo DoDesequipar_Err
        '</EhHeader>

        Dim Probabilidad   As Integer

        Dim Resultado      As Integer

        Dim WrestlingSkill As Byte

        Dim AlgoEquipado   As Boolean
    
100     With UserList(UserIndex)

            ' Si no tiene guantes de hurto no desequipa.
102         If .Invent.WeaponEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
            ' Si no esta solo con manos, no desequipa tampoco.
104         If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
106         WrestlingSkill = .Stats.UserSkills(eSkill.Armas)
        
108         Probabilidad = WrestlingSkill * 0.2 + (.Stats.Elv) * 0.66
        End With
   
110     With UserList(VictimIndex)

            ' Si tiene escudo, intenta desequiparlo
112         If .Invent.EscudoEqpObjIndex > 0 Then
            
114             Resultado = RandomNumber(1, 100)
            
116             If Resultado <= Probabilidad Then
                    ' Se lo desequipo
118                 Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
120                 Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
122                 Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                
124                 Call FlushBuffer(VictimIndex)
                
                    Exit Sub

                End If
            
126             AlgoEquipado = True
            End If
        
            ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
128         If .Invent.WeaponEqpObjIndex > 0 Then
            
130             Resultado = RandomNumber(1, 100)
            
132             If Resultado <= Probabilidad Then
                    ' Se lo desequipo
134                 Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
136                 Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
138                 Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
                
140                 Call FlushBuffer(VictimIndex)
                
                    Exit Sub

                End If
            
142             AlgoEquipado = True
            End If
        
            ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
144         If .Invent.CascoEqpObjIndex > 0 Then
            
146             Resultado = RandomNumber(1, 100)
            
148             If Resultado <= Probabilidad Then
                    ' Se lo desequipo
150                 Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
152                 Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
154                 Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
                
156                 Call FlushBuffer(VictimIndex)
                
                    Exit Sub

                End If
            
158             AlgoEquipado = True
            End If
    
160         If AlgoEquipado Then
162             Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
            Else
164             Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            End If
    
        End With

        '<EhFooter>
        Exit Sub

DoDesequipar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoDesequipar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
        '<EhHeader>
        On Error GoTo DoHurtar_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modif: 03/03/2010
        'Implements the pick pocket skill of the Bandit :)
        '03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
        '***************************************************
        Dim OtroUserIndex As Integer

100     If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

        Exit Sub

        'If UserList(UserIndex).Clase <> eClass.Bandit Then Exit Sub
        'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
        'Uso el slot de los anillos para "equipar" los guantes.
        'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
102     If UserList(UserIndex).Invent.WeaponEqpObjIndex <> GUANTE_HURTO Then Exit Sub

        Dim res As Integer

104     res = RandomNumber(1, 100)

106     If (res < 20) Then
108         If TieneObjetosRobables(VictimaIndex) Then
    
110             If UserList(VictimaIndex).flags.Comerciando Then
112                 OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
114                 If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
116                     Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
118                     Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
120                     Call LimpiarComercioSeguro(VictimaIndex)
122                     Call Protocol.FlushBuffer(OtroUserIndex)
                    End If
                End If
                
124             Call RobarObjeto(UserIndex, VictimaIndex)
126             Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
            Else
128             Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        '<EhFooter>
        Exit Sub

DoHurtar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoHurtar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
        '<EhHeader>
        On Error GoTo DoHandInmo_Err
        '</EhHeader>

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modif: 17/02/2007
        'Implements the special Skill of the Thief
        '***************************************************
100     If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
102     If UserList(UserIndex).Clase <> eClass.Thief Then Exit Sub
    
104     If UserList(UserIndex).Invent.WeaponEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        Dim res As Integer

106     res = RandomNumber(0, 100)

108     If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) / 4) Then
110         UserList(VictimaIndex).flags.Paralizado = 1
112         UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
        
114         UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
116         UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).Name
        
118         Call WriteParalizeOK(VictimaIndex)
120         Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
122         Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
        End If

        '<EhFooter>
        Exit Sub

DoHandInmo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.DoHandInmo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/04/2010 (ZaMa)
        '02/04/2010: ZaMa - Nueva formula para desarmar.
        '***************************************************
        '<EhHeader>
        On Error GoTo Desarmar_Err
        '</EhHeader>

        Dim Probabilidad   As Integer

        Dim Resultado      As Integer

        Dim WrestlingSkill As Byte
    
100     With UserList(UserIndex)
102         WrestlingSkill = .Stats.UserSkills(eSkill.Armas)
        
104         Probabilidad = WrestlingSkill * 0.2 + (.Stats.Elv) * 0.66
        
106         Resultado = RandomNumber(1, 100)
        
108         If Resultado <= Probabilidad Then
110             Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
112             Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
114             Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

116             Call FlushBuffer(VictimIndex)
            End If

        End With
    
        '<EhFooter>
        Exit Sub

Desarmar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.Desarmar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/01/2010
        '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
        '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
        '***************************************************
        '<EhHeader>
        On Error GoTo MaxItemsConstruibles_Err
        '</EhHeader>
    
100     With UserList(UserIndex)

104            MaxItemsConstruibles = MaximoInt(1, CInt(((.Stats.Elv) - 2) * 0.2))


        End With

        '<EhFooter>
        Exit Function

MaxItemsConstruibles_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.MaxItemsConstruibles " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/05/2010
        '***************************************************
        '<EhHeader>
        On Error GoTo MaxItemsExtraibles_Err
        '</EhHeader>
100     MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1
        '<EhFooter>
        Exit Function

MaxItemsExtraibles_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.MaxItemsExtraibles " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub ImitateNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        'Copies body, head and desc from previously clicked npc.
        '***************************************************
        '<EhHeader>
        On Error GoTo ImitateNpc_Err
        '</EhHeader>
    
100     With UserList(UserIndex)
        
            ' Copy desc
102         .DescRM = Npclist(NpcIndex).Name
        
            ' Remove Anims (Npcs don't use equipment anims yet)
104         .Char.CascoAnim = NingunCasco
106         .Char.ShieldAnim = NingunEscudo
108         .Char.WeaponAnim = NingunArma
        
            ' If admin is invisible the store it in old char
110         If .flags.AdminInvisible = 1 Or .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            
112             .flags.OldBody = Npclist(NpcIndex).Char.Body
114             .flags.OldHead = Npclist(NpcIndex).Char.Head
            Else
116             .Char.Body = Npclist(NpcIndex).Char.Body
118             .Char.Head = Npclist(NpcIndex).Char.Head
            
120             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
            End If
    
        End With
    
        '<EhFooter>
        Exit Sub

ImitateNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Trabajo.ImitateNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

