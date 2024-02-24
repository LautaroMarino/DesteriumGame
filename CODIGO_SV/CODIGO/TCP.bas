Attribute VB_Name = "TCP"
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

Function IBody_Generate(ByVal UserGenero As Byte, ByVal UserRaza As Byte) As Integer
        '<EhHeader>
        On Error GoTo DarCuerpoYCabeza_Err
        '</EhHeader>

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 14/03/2007
        'Elije una cabeza para el usuario y le da un body
        '*************************************************
        Dim NewBody    As Integer

        Dim NewHead    As Integer

    
104     Select Case UserGenero

            Case eGenero.Hombre

106             Select Case UserRaza

                    Case eRaza.Humano
108                     NewHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)
110                     NewBody = HUMANO_H_CUERPO_DESNUDO

112                 Case eRaza.Elfo
114                     NewHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)
116                     NewBody = ELFO_H_CUERPO_DESNUDO

118                 Case eRaza.Drow
120                     NewHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)
122                     NewBody = DROW_H_CUERPO_DESNUDO

124                 Case eRaza.Enano
126                     NewHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)
128                     NewBody = ENANO_H_CUERPO_DESNUDO

130                 Case eRaza.Gnomo
132                     NewHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)
134                     NewBody = GNOMO_H_CUERPO_DESNUDO
                End Select

136         Case eGenero.Mujer

138             Select Case UserRaza

                    Case eRaza.Humano
140                     NewHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)
142                     NewBody = HUMANO_M_CUERPO_DESNUDO

144                 Case eRaza.Elfo
146                     NewHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)
148                     NewBody = ELFO_M_CUERPO_DESNUDO

150                 Case eRaza.Drow
152                     NewHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)
154                     NewBody = DROW_M_CUERPO_DESNUDO

156                 Case eRaza.Enano
158                     NewHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)
160                     NewBody = ENANO_M_CUERPO_DESNUDO

162                 Case eRaza.Gnomo
164                     NewHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)
166                     NewBody = GNOMO_M_CUERPO_DESNUDO
                End Select
        End Select
    
        IBody_Generate = NewBody
        '<EhFooter>
        Exit Function

DarCuerpoYCabeza_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.DarCuerpoYCabeza " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function IHead_Generate(ByVal UserGenero As Byte, ByVal UserRaza As Byte) As Integer
        '<EhHeader>
        On Error GoTo DarCabezaRandom_Err
        '</EhHeader>
        Dim NewBody    As Integer

        Dim NewHead    As Integer
    
104     Select Case UserGenero

            Case eGenero.Hombre

106             Select Case UserRaza

                    Case eRaza.Humano
108                     NewHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)

110                 Case eRaza.Elfo
112                     NewHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)

114                 Case eRaza.Drow
116                     NewHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)

118                 Case eRaza.Enano
120                     NewHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)

122                 Case eRaza.Gnomo
124                     NewHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)
                End Select

126         Case eGenero.Mujer

128             Select Case UserRaza

                    Case eRaza.Humano
130                     NewHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)


132                 Case eRaza.Elfo
134                     NewHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)


136                 Case eRaza.Drow
138                     NewHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)


140                 Case eRaza.Enano
142                     NewHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)

144                 Case eRaza.Gnomo
146                     NewHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)

                End Select
        End Select
    
150    IHead_Generate = NewHead
        '<EhFooter>
        Exit Function

DarCabezaRandom_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.DarCabezaRandom " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Function ValidarCabeza(ByVal UserRaza As Byte, _
                       ByVal UserGenero As Byte, _
                       ByVal Head As Integer) As Boolean
        '<EhHeader>
        On Error GoTo ValidarCabeza_Err
        '</EhHeader>

100     Select Case UserGenero

            Case eGenero.Hombre

102             Select Case UserRaza

                    Case eRaza.Humano
104                     ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And Head <= HUMANO_H_ULTIMA_CABEZA)

106                 Case eRaza.Elfo
108                     ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And Head <= ELFO_H_ULTIMA_CABEZA)

110                 Case eRaza.Drow
112                     ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And Head <= DROW_H_ULTIMA_CABEZA)

114                 Case eRaza.Enano
116                     ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And Head <= ENANO_H_ULTIMA_CABEZA)

118                 Case eRaza.Gnomo
120                     ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And Head <= GNOMO_H_ULTIMA_CABEZA)
                End Select
    
122         Case eGenero.Mujer

124             Select Case UserRaza

                    Case eRaza.Humano
126                     ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And Head <= HUMANO_M_ULTIMA_CABEZA)

128                 Case eRaza.Elfo
130                     ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And Head <= ELFO_M_ULTIMA_CABEZA)

132                 Case eRaza.Drow
134                     ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And Head <= DROW_M_ULTIMA_CABEZA)

136                 Case eRaza.Enano
138                     ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And Head <= ENANO_M_ULTIMA_CABEZA)

140                 Case eRaza.Gnomo
142                     ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And Head <= GNOMO_M_ULTIMA_CABEZA)
                End Select
        End Select
        
        '<EhFooter>
        Exit Function

ValidarCabeza_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ValidarCabeza " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function AsciiValidos(ByVal cad As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo AsciiValidos_Err
        '</EhHeader>

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
          
106         If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
108             AsciiValidos = False

                Exit Function

            End If
          
110     Next i

112     AsciiValidos = True

        '<EhFooter>
        Exit Function

AsciiValidos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.AsciiValidos " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Function ValidarNombre(Nombre As String) As Boolean
        '<EhHeader>
        On Error GoTo ValidarNombre_Err
        '</EhHeader>
    
100     If Len(Nombre) < ACCOUNT_MIN_CHARACTER_CHAR Or Len(Nombre) > ACCOUNT_MAX_CHARACTER_CHAR Then Exit Function
    
        Dim Temp As String, CantidadEspacios As Byte
102     Temp = UCase$(Nombre)
    
        Dim i As Long, Char As Integer, LastChar As Integer
104     For i = 1 To Len(Temp)
106         Char = Asc(mid$(Temp, i, 1))
        
108         If (Char < 65 Or Char > 90) And Char <> 32 Then
                Exit Function
        
110         ElseIf Char = 32 Then

112             If LastChar = 32 Then
                    Exit Function
                End If
                
114             CantidadEspacios = CantidadEspacios + 1
                
116             If CantidadEspacios > 1 Then
                    Exit Function
                End If
            End If
        
118         LastChar = Char
        Next

120     If Asc(mid$(Temp, 1, 1)) = 32 Or Asc(mid$(Temp, Len(Temp), 1)) = 32 Then
            Exit Function
        End If
    
122     ValidarNombre = True

        '<EhFooter>
        Exit Function

ValidarNombre_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ValidarNombre " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
Function Numeric(ByVal cad As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo Numeric_Err
        '</EhHeader>

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If (car < 48 Or car > 57) Then
108             Numeric = False

                Exit Function

            End If
    
110     Next i

112     Numeric = True

        '<EhFooter>
        Exit Function

Numeric_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.Numeric " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo NombrePermitido_Err
        '</EhHeader>

        Dim i As Integer

100     For i = 1 To UBound(ForbidenNames)

102         If InStr(Nombre, ForbidenNames(i)) Then
104             NombrePermitido = False

                Exit Function

            End If

106     Next i

108     NombrePermitido = True

        '<EhFooter>
        Exit Function

NombrePermitido_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.NombrePermitido " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function PalabraPermitida(ByVal Texto As String) As Boolean
        ' Realiza una comparación de palabras permitidas, previamente sacamos los espacios.
        ' Pasar textos siempre con lcase$()
        '<EhHeader>
        On Error GoTo PalabraPermitida_Err
        '</EhHeader>
    
        Dim i As Integer
    
100     Texto = Replace(Texto, " ", "")
102     For i = 1 To UBound(ForbidenText)

104         If InStr(Texto, ForbidenText(i)) Then
106             PalabraPermitida = False

                Exit Function

            End If

108     Next i

110     PalabraPermitida = True

        '<EhFooter>
        Exit Function

PalabraPermitida_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.PalabraPermitida " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function AsciiValidos_Chat(ByVal cad As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo AsciiValidos_Chat_Err
        '</EhHeader>

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
          
106         If (car = 126) Then
108             AsciiValidos_Chat = False

                Exit Function

            End If
          
110     Next i

112     AsciiValidos_Chat = True

        '<EhFooter>
        Exit Function

AsciiValidos_Chat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.AsciiValidos_Chat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ValidateSkills_Err
        '</EhHeader>

        Dim LoopC As Integer

100    ' For LoopC = 1 To NUMSKILLS

102       '  If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then

                'Exit Function

104          '   If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
          '  End If

106   '  Next LoopC

108     ValidateSkills = True
    
        '<EhFooter>
        Exit Function

ValidateSkills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ValidateSkills " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function ConnectNewUser(ByVal UserName As String, _
                   ByVal UserClase As eClass, _
                   ByVal UserRaza As eRaza, _
                   ByVal UserSexo As eGenero, _
                   ByVal UserHead As Integer, _
                   ByRef IUser As User) As User

        On Error GoTo ConnectNewUser_Err

        Dim i As Long
        
         Call ResetUserFlags(IUser)
         
100     With IUser
              .LastHeading = 0
              .flags.Privilegios = 0
102         .flags.TargetX = 0
104         .flags.TargetY = 0
106         .flags.TargetMap = 0
        
110         .Stats.Elv = 1
112         .Stats.Exp = 0

114         .Stats.Elu = 300
              
              
116         .Stats.Gld = 0
118         .Stats.Eldhir = 0

            .Stats.BonusLast = 0
            ReDim .Stats.Bonus(0) As UserBonus
124         '.Stats.Retos1Ganados = 0
126         '.Stats.DesafiosGanados = 0
128         '.Stats.Retos1Jugados = 0
130         '.Stats.DesafiosJugados = 0
132         '.Stats.TorneosGanados = 0
134         '.Stats.TorneosJugados = 0
136         .flags.SlotRetoUser = 255
138         .flags.Muerto = 0
140         .flags.Escondido = 0
    
142         .Reputacion.AsesinoRep = 0
144         .Reputacion.BandidoRep = 0
146         .Reputacion.BurguesRep = 0
148         .Reputacion.LadronesRep = 0
150         .Reputacion.NobleRep = 1000
152         .Reputacion.PlebeRep = 30
    
154         .Reputacion.promedio = 30 / 6
    
156         .Name = UserName
158         .Clase = UserClase
160         .Raza = UserRaza
162         .Genero = UserSexo
164         .Hogar = cUllathorpe
        
            ' Dados 18
166         .Stats.UserAtributos(eAtributos.Fuerza) = 18 + Balance.ModRaza(UserRaza).Fuerza
168         .Stats.UserAtributos(eAtributos.Agilidad) = 18 + Balance.ModRaza(UserRaza).Agilidad
170         .Stats.UserAtributos(eAtributos.Inteligencia) = 18 + Balance.ModRaza(UserRaza).Inteligencia
172         .Stats.UserAtributos(eAtributos.Carisma) = 18 + Balance.ModRaza(UserRaza).Carisma
174         .Stats.UserAtributos(eAtributos.Constitucion) = 18 + Balance.ModRaza(UserRaza).Constitucion
    
186         .Char.Heading = eHeading.SOUTH
        
188         If .Account.Premium > 0 Then
190             .Char.Head = UserHead
            Else
                 .Char.Head = IHead_Generate(UserSexo, UserRaza)
            End If
                
              .Char.Body = IBody_Generate(UserSexo, UserRaza)
              .Char.ShieldAnim = NingunEscudo
              .Char.CascoAnim = NingunCasco
              .Char.WeaponAnim = NingunArma
194         .OrigChar = .Char
    
            #If ConUpTime Then
196             .LogOnTime = Now
198             .UpTime = 0
            #End If

        
        End With
        
        ConnectNewUser = IUser
        
206

        
        '<EhFooter>
        Exit Function

ConnectNewUser_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ConnectNewUser " & "at line " & Erl

        

        '</EhFooter>
End Function

Sub CloseSocket(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CloseSocket_Err
        '</EhHeader>

        Dim isNotVisible As Boolean

        Dim HiddenPirat  As Boolean
    
100     With UserList(UserIndex)
102         isNotVisible = (.flags.Oculto Or .flags.Invisible)

104         If isNotVisible Then
106             .flags.Invisible = 0
108             .flags.Oculto = 0
                
                ' Para no repetir mensajes
110             Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si esta navegando ya esta visible
112             If .flags.Navegando = 0 Then
114                 Call SetInvisible(UserIndex, .Char.charindex, False)
                End If
            End If
            
116         If .flags.Traveling = 1 Then
118             Call EndTravel(UserIndex, True)
            End If
        
            'mato los comercios seguros
120         If .ComUsu.DestUsu > 0 Then
122             If UserList(.ComUsu.DestUsu).flags.UserLogged Then
124                 If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
126                     Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
128                     Call FinComerciarUsu(.ComUsu.DestUsu)
                    End If
                End If
            End If
        
            ' Eventos automáticos
            Dim SlotEvent As Byte, SlotUser As Byte, TeamUser As Byte, MapFight As Byte
130         SlotEvent = .flags.SlotEvent
132         SlotUser = .flags.SlotUserEvent
        
134         If SlotEvent > 0 Then
136             TeamUser = Events(SlotEvent).Users(SlotUser).Team
138             MapFight = Events(SlotEvent).Users(SlotUser).MapFight
            
140             Call AbandonateEvent(UserIndex, , True)
            
                ' Si no empezo no tiene sentido comprobar esto. Es para buscar ganador
142             If Events(SlotEvent).Run Then Call Events_CheckInscribed(UserIndex, SlotEvent, SlotUser, TeamUser, MapFight)
            End If
              
            ' Retos entre personajes
144         If .flags.SlotReto > 0 Then
146             Call mRetos.UserdieFight(UserIndex, 0, True)
            End If
        
148         If .flags.Desafiando > 0 Then
150             Desafio_UserKill UserIndex
            End If
        
152         If .flags.SlotFast > 0 Then
154             RetoFast_UserDie UserIndex, True
            End If
        
156         If .flags.Transform Then
158             Call Transform_User(UserIndex, 0)
            End If
        
160         If .flags.TransformVIP Then
162             Call TransformVIP_User(UserIndex, 0)
            End If
        
164         If .flags.ClainObject = 1 Then
166             Call mRetos.Retos_ReclameObj(UserIndex)
            End If
        
168         If Power.UserIndex = UserIndex Then
170             Call Power_Set(0, UserIndex)
            End If
            
            If .flags.BotList > 0 Then
                Call Streamer_SetBotList(UserIndex, .flags.BotList, True)
            End If
            
           Call Teleports_Cancel(UserIndex)
           
            If Not EsGm(UserIndex) Then
                Call WriteUpdateUserData(UserList(UserIndex))
            End If
            
172         If .Pos.Map > 0 Then
                
174             If .GuildIndex > 0 Then
176                 GuildsInfo(.GuildIndex).Members(.GuildSlot).UserIndex = 0
178                 .GuildSlot = 0
180                 Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("El personaje " & .Name & " se ha desconectado.", FontTypeNames.FONTTYPE_GUILDMSG))
            
                End If
            
182             If MapInfo(.Pos.Map).LvlMin > .Stats.Elv Then
184                 Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, False)
                End If
            
            End If

              If StreamerBot.Active = UserIndex Then
                    Call Streamer_Initial(0, 0, 0, 0)
              End If
              
              'Call Streamer_CheckUser(UserIndex)
              
186         If .flags.UserLogged Then
188             Call CloseUser(UserIndex)

          '  #If Classic = 0 Then
               ' Battle_Arenas(.ServerSelected).Users = Battle_Arenas(.ServerSelected).Users - 1
             '   .ServerSelected = 0
           '   WriteLoggedAccountBattle UserIndex
            '  #End If
              
            End If
        
190         Call ResetUserSlot(UserIndex)
            

        End With


   
        '<EhFooter>
        Exit Sub

CloseSocket_Err:
         Call ResetUserSlot(UserIndex)

        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.CloseSocket " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
        '<EhHeader>
        On Error GoTo EstaPCarea_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        Dim X As Integer, Y As Integer

100     For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
102         For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

104             If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
106                 EstaPCarea = True

                    Exit Function

                End If
        
108         Next X
110     Next Y

112     EstaPCarea = False
        '<EhFooter>
        Exit Function

EstaPCarea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.EstaPCarea " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo HayPCarea_Err
        '</EhHeader>

        Dim X As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If X > 0 And Y > 0 And X < 101 And Y < 101 Then
106                 If MapData(Pos.Map, X, Y).UserIndex > 0 Then
108                     HayPCarea = True

                        Exit Function

                    End If
                End If

110         Next X
112     Next Y

114     HayPCarea = False
        '<EhFooter>
        Exit Function

HayPCarea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.HayPCarea " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo HayOBJarea_Err
        '</EhHeader>

        Dim X As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
106                 HayOBJarea = True

                    Exit Function

                End If
        
108         Next X
110     Next Y

112     HayOBJarea = False
        '<EhFooter>
        Exit Function

HayOBJarea_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.HayOBJarea " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ValidateChr_Err
        '</EhHeader>

100     ValidateChr = UserList(UserIndex).Char.Head <> 0 And UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

        '<EhFooter>
        Exit Function

ValidateChr_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ValidateChr " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function CheckPenas(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckPenas_Err
        '</EhHeader>
 
    Dim tStr As String
    
100         If val(GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "BAN")) > 0 Then
            
102             tStr = GetVar(CharPath & UCase$(Name) & ".chr", "PENAS", "DATEDAY")
            
104             If tStr <> vbNullString Then
106                 If Format(Now, "dd/mm/yyyy") > tStr Then
108                     Call UnBan(UCase$(Name))
                    End If
                
                Else

                    Dim Razon As String

110                 Dim Pena  As String: Pena = GetVar(CharPath & UCase$(Name) & ".chr", "PENAS", "CANT")

112                 Razon = GetVar(CharPath & UCase$(Name) & ".chr", "PENAS", "P" & Pena)
114                 Call WriteErrorMsg(UserIndex, "Tu personaje no tiene permitido ingresar al juego. RAZON: " & Razon)

                    Exit Function

                End If
           
            End If
            
116         CheckPenas = True
        '<EhFooter>
        Exit Function

CheckPenas_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.CheckPenas " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub UpdatePremium(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo UpdatePremium_Err

        '</EhHeader>

        Dim TimerPremium As String

100     TimerPremium = UserList(UserIndex).Account.DatePremium

102     If TimerPremium <> vbNullString Then
            If DateDiff("s", Now, TimerPremium) <= 0 Then
104             ' If Format(Now, "dd/mm/aa hh:mm:ss") > TimerPremium Then
106             UserList(UserIndex).Account.DatePremium = vbNullString
108             UserList(UserIndex).Account.Premium = 0
110             Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "INIT", "DATEPREMIUM", vbNullString)
112             Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "INIT", "PREMIUM", "0")
            
114             Call WriteConsoleMsg(UserIndex, "¡El PREMIUM se ha ido de tu cuenta!", FontTypeNames.FONTTYPE_INFORED)
            
            Else
                Call WriteConsoleMsg(UserIndex, "Tu cuenta PREMIUM vence " & UserList(UserIndex).Account.DatePremium & ".", FontTypeNames.FONTTYPE_INFOGREEN)
            End If

        End If

        '<EhFooter>
        Exit Sub

UpdatePremium_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.UpdatePremium " & "at line " & Erl

        

        '</EhFooter>
End Sub

' # Chequea si ya pasó la hora y son llevados a donde empieza la versión
Public Sub User_Go_Initial_Version()

    If Not DateDiff("s", Now, DateAperture) <= 0 Then Exit Sub
    
    Dim A As Long
    Dim Pos As WorldPos
    
    For A = 1 To LastUser
        With UserList(A)
            If .Pos.Map = CiudadFlotante.Map Then
                Pos.Map = Newbie.Map
                Pos.X = RandomNumber(Newbie.X - 5, Newbie.X + 5)
                Pos.Y = RandomNumber(Newbie.Y - 5, Newbie.Y + 5)
                
                Call EventWarpUser(A, Pos.Map, Pos.X, Pos.Y)
                 Call SendData(SendTarget.ToOne, A, PrepareMessagePlayEffect(eSound.sVictory4, Pos.X, Pos.Y))
                 
                Call WriteConsoleMsg(A, "¡La espera terminó! A entrenar y disfrutar de una nueva versión", FontTypeNames.FONTTYPE_DESAFIOS)
                
            End If
        End With
    Next A
    
   
End Sub
' # Chequea el momento en el que logea si está en previa de apertura o ya comenzó la versión
Public Function User_Check_Login_Apertura(ByVal UserIndex As Integer) As WorldPos

    Dim Pos As WorldPos
    
    With UserList(UserIndex)
        
        
        If .Pos.Map = 0 Then
            ' # Ya comenzó la versión
            
                If DateDiff("s", Now, DateAperture) <= 0 Then
                    Pos.Map = Newbie.Map
                    Pos.Y = Newbie.Y
                    Pos.X = RandomNumber(Newbie.X - 3, Newbie.X + 1)
                Else
                    ' # Sum en ciudad flotante
                    Pos.Map = CiudadFlotante.Map
                    Pos.Y = CiudadFlotante.Y
                    Pos.X = RandomNumber(CiudadFlotante.X - 3, CiudadFlotante.X + 3)
                End If
            
        Else
            ' El personaje deslogeo antes de tiempo y quedo en la flotante
            If .Pos.Map = CiudadFlotante.Map Then
                ' # Ya comenzó la versión. Lo llevamos al dungeon newbie
                If DateDiff("s", Now, DateAperture) <= 0 Then
                    Pos.Map = Newbie.Map
                    Pos.Y = Newbie.Y
                    Pos.X = RandomNumber(Newbie.X - 3, Newbie.X + 1)
                End If
            End If
            
        End If
    End With
    
    User_Check_Login_Apertura = Pos
End Function
Sub ConnectUser(ByVal UserIndex As Integer, ByRef Name As String, ByVal NewChar As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 24/07/2010 (ZaMa)
    '26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
    '12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
    '14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
    '11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
    '03/12/2009: Budi - Optimización del código
    '24/07/2010: ZaMa - La posicion de comienzo es namehuak, como se habia definido inicialmente.
    '***************************************************


    On Error GoTo ErrHandler
    
    Dim N     As Integer
        
    Dim A As Long
        

    Dim tStr  As String

    Dim Valid As Boolean

    With UserList(UserIndex)
    
        Dim i As Long
         
        Call ResetUserFlags(UserList(UserIndex))
        
       If Not CheckPenas(UserIndex, Name) Then Exit Sub

         '¿Ya esta conectado el personaje?
        If CheckForSameName(Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
            Else
                Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")
            End If

            Exit Sub

        End If
            
        If val(GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "BLOCKED")) > 0 Then
              Dim TempOfferTime As String
              TempOfferTime = GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "OFFERTIME")
                  
              If TempOfferTime <> vbNullString Then
                     If Format(Now, "dd/mm/aa hh:mm:ss") > TempOfferTime Then
                        Call WriteVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "OFFERTIME", "")
                        Call WriteVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "BLOCKED", "0")
                     Else
                        Call WriteErrorMsg(UserIndex, "Tu personaje ha sido ofrecido en MODO CANDADO. Esto significa que podrás entrar: " & TempOfferTime)
                     End If
                         
              Else
                Call WriteErrorMsg(UserIndex, "El personaje está bloqueado ya que está en el mercado central. Deberás quitarlo de la misma para poder ingresar.")
                  Exit Sub
              End If
        End If
            
          'Reseteamos los FLAGS
        .UserKey = 0
        .UserLastClick = 0
        .UserLastClick_Tolerance = 0
    
        .flags.Escondido = 0
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Char.FX = 0
        .flags.MenuCliente = 255
        .flags.LastSlotClient = 255
    
           
    
            
        'Reseteamos los privilegios
        .flags.Privilegios = 0
    
        'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
        If EsAdmin(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call Logs_User(Name, eLog.eGm, eLogDescUser.eNone, "Se conecto con ip:" & .IpAddress)
        ElseIf EsDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call Logs_User(Name, eLog.eGm, eLogDescUser.eNone, "Se conecto con ip:" & .IpAddress)
        ElseIf EsSemiDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        
            .flags.PrivEspecial = EsGmEspecial(Name)
        
            Call Logs_User(Name, eLog.eGm, eLogDescUser.eNone, "Se conecto con ip:" & .IpAddress)
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True
        End If
        
        If ServerSoloGMs > 0 Then
            If Not Email_Is_Testing_Pro(.Account.Email) Then
                Call Protocol.Kick(UserIndex, "Servidor en mantenimiento. Consulta otros servidores para disfrutar y pasar el rato.")
        
                Exit Sub
        
            End If
        End If

        'Cargamos el personaje
        Dim Leer As clsIniManager

        Set Leer = New clsIniManager

        Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")
        
        
        ' Cargamos la reputación antes para generar algunos cambios sobre los flags
        Call LoadUserReputacion(UserIndex, Leer)
        
        'Cargamos los datos del personaje
        Call LoadUserInit(UserIndex, Leer)

        Call LoadUserStats(UserIndex, Leer)

        Call LoadQuestStats(UserIndex, Leer)

        Call LoadUserAntiFrags(UserIndex, Leer)

        'Cargamos los mensajes privados del usuario.
        Call CargarMensajes(UserIndex, Leer)
    
        If Not ValidateChr(UserIndex) Then
            Call Protocol.Kick(UserIndex, "Error en el personaje.")

            Exit Sub

        End If
    
        
        Call LoadUserMeditations(UserIndex, Leer)

        Set Leer = Nothing
              
        If .Invent.ArmourEqpObjIndex > 0 Then .Char.AuraIndex(1) = ObjData(.Invent.ArmourEqpObjIndex).AuraIndex(1)
        If .Invent.WeaponEqpObjIndex > 0 Then .Char.AuraIndex(2) = ObjData(.Invent.WeaponEqpObjIndex).AuraIndex(2)
        If .Invent.CascoEqpObjIndex > 0 Then .Char.AuraIndex(3) = ObjData(.Invent.CascoEqpObjIndex).AuraIndex(3)
        If .Invent.EscudoEqpObjIndex > 0 Then .Char.AuraIndex(4) = ObjData(.Invent.EscudoEqpObjIndex).AuraIndex(4)
        
        If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
              
        If .Invent.MochilaEqpSlot > 0 Then
            .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).ObjIndex).MochilaType * 5
        Else
            .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End If

        .flags.DragBlocked = False

        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)

        If .flags.Paralizado Then
            Call WriteParalizeOK(UserIndex)
        End If

        Dim mapa          As Integer

        Dim MessageNewbie As String
        
        mapa = .Pos.Map
        
        Dim TempPos As WorldPos
        TempPos = User_Check_Login_Apertura(UserIndex)
        
        If TempPos.X <> 0 Then
            .Pos = TempPos
        End If
        
        'Posicion de comienzo
        If mapa = 0 Then
            mapa = .Pos.Map
            
        Else
             ' El personaje deslogeo antes de tiempo y quedo en la flotante
            If mapa = CiudadFlotante.Map Then
                ' # Ya comenzó la versión. Lo llevamos al dungeon newbie
                If DateDiff("s", Now, DateAperture) <= 0 Then
                    .Pos.Map = Newbie.Map
                    .Pos.Y = Newbie.Y
                    .Pos.X = RandomNumber(Newbie.X - 3, Newbie.X + 1)
                End If
            End If
            
            If Not MapaValido(mapa) Then
                Call Protocol.Kick(UserIndex, "El PJ se encuenta en un mapa inválido.")

                Exit Sub

            End If
        
            ' If map has different initial coords, update it
            Dim StartMap As Integer

            StartMap = MapInfo(mapa).StartPos.Map

            If StartMap <> 0 Then
                If MapaValido(StartMap) Then
                    .Pos = MapInfo(mapa).StartPos
                    mapa = StartMap
                End If
            End If
        
        End If
    
        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
        If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

            Dim FoundPlace As Boolean

            Dim esAgua     As Boolean

            Dim tX         As Long

            Dim tY         As Long
        
            FoundPlace = False
            esAgua = HayAgua(mapa, .Pos.X, .Pos.Y)
        
            For tY = .Pos.Y - 1 To .Pos.Y + 1
                For tX = .Pos.X - 1 To .Pos.X + 1

                    If esAgua Then

                        'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                        If LegalPos(mapa, tX, tY, True, False, True) Then
                            FoundPlace = True

                            Exit For

                        End If

                    Else

                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                        If LegalPos(mapa, tX, tY, False, True, True) Then
                            FoundPlace = True

                            Exit For

                        End If
                    End If

                Next tX
            
                If FoundPlace Then Exit For
            Next tY
        
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.X = tX
                .Pos.Y = tY
            Else

                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                    If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                        'Le avisamos al que estaba comerciando que se tuvo que ir.
                        If UserList(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                            Call FinComerciarUsu(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                            Call WriteConsoleMsg(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                            Call FlushBuffer(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                        End If
                    
                        'Lo sacamos.
                        If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                            Call FinComerciarUsu(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                        End If
                    End If
                
                    If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                        Call WriteErrorMsg(MapData(mapa, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                    End If
                
                    'Call CloseSocket(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)
                    Call WriteDisconnect(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                    Call FlushBuffer(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                    Call CloseSocket(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                End If
            End If
        End If

        'Nombre de sistema
        .Name = Name
        .secName = .Name
    
        .ShowName = True 'Por default los nombres son visibles
    
        'If in the water, and has a boat, equip it!
        If .Invent.BarcoObjIndex > 0 And (HayAgua(mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.Body)) Then

            .Char.Head = 0

            If .flags.Muerto = 0 Then
                Call ToggleBoatBody(UserIndex)
            Else
                .Char.Body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
                  For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = NingunAura
                  Next A
            End If
        
            .flags.Navegando = 1
        End If
    
        'Info
        Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        Call WriteChangeMap(UserIndex, .Pos.Map) 'Carga el mapa
        Call WritePlayMusic(UserIndex, val(ReadField(1, MapInfo(.Pos.Map).Music, 45)))

        If .flags.Privilegios = PlayerType.Dios Then
            .flags.ChatColor = RGB(250, 250, 150)
        ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 0)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 255)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
            .flags.ChatColor = RGB(255, 128, 64)
        Else
            .flags.ChatColor = vbWhite
        End If

        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
    
        'Crea  el personaje del usuario
        If Not MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y) Then
            Exit Sub
        End If

        If (.flags.Privilegios And (PlayerType.User)) = 0 Then
            Call DoAdminInvisible(UserIndex)
            .flags.SendDenounces = True
        Else

            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
            End If
        End If

        Call WriteUserCharIndexInServer(UserIndex)
        Call ActualizarVelocidadDeUsuario(UserIndex, False)
        ''[/el oso]

        ' // NUEVO
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then
            Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
        End If
        
        Call CheckUserLevel(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
    
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)

        'Call SendMOTD(UserIndex)

        If haciendoBK Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin)
        End If

        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin)
        End If

        If EnTesting Then
            Call WriteErrorMsg(UserIndex, "Servidor en Testeo. Espere unos momentos y consulte la página oficial. WWW.ARGENTUMGAME.COM")

            Exit Sub

        End If

        If TieneMensajesNuevos(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "¡Tienes mensajes privados sin leer!", FontTypeNames.FONTTYPE_FIGHT)
        End If

        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
        
        .flags.UserLogged = True

        'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        MapInfo(.Pos.Map).Players.Add UserIndex
     
        If .Stats.SkillPts > 0 Then
            Call WriteLevelUp(UserIndex, .Stats.SkillPts)
        End If

        If NumUsers + UsersBot > RECORDusuarios Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Seguimos sumando jugadores a nuestra comunidad!." & " Hay " & NumUsers + UsersBot & " usuarios conectados. Gracias por Jugar.", FontTypeNames.FONTTYPE_INFO))
            RECORDusuarios = NumUsers + UsersBot
            Call WriteVar(IniPath & "Server.ini", "INIT", "RECORD", Str(RECORDusuarios))
        
            'Call EstadisticasWeb.Informar(RECORD_USUARIOS, RECORDusuarios)
        End If

        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(UserIndex)
        End If

        If .flags.Montando = 1 Then
            Call WriteMontateToggle(UserIndex)
        End If
        
        'If .flags.Muerto = 1 Then
            Call WriteUpdateUserDead(UserIndex, .flags.Muerto)
        
        'End If
        
        Call WriteConsoleMsg(UserIndex, "Desterium Online. Un servidor de Argentum Online.", FontTypeNames.FONTTYPE_CONSEJOVesA)
       ' Call WriteConsoleMsg(UserIndex, "Utiliza el comando /AYUDA. ¡Te dirá todo lo que necesitas saber para comenzar! Recuerda que desde la página principal podrás acceder a soporte 24/7", FontTypeNames.FONTTYPE_USERGOLD)
        
        If HappyHour Then
            Call WriteConsoleMsg(UserIndex, "¡HappyHour Activado! Exp x2 ¡Entrená tu personaje!", FontTypeNames.FONTTYPE_USERBRONCE)
        End If
            
        If PartyTime Then
            Call WriteConsoleMsg(UserIndex, "PartyTime» Los miembros de la party reciben 25% de experiencia extra.", FontTypeNames.FONTTYPE_INVASION)
        End If
            
        'Call WriteConsoleMsg(UserIndex, "MANUAL: WWW.ARGENTUMGAME.COM/wiki/", FontTypeNames.FONTTYPE_USERBRONCE)
        Call WriteConsoleMsg(UserIndex, "Cualquier acto considerado dañino para la comunidad y/o usuarios miembros de la misma retornará en un bloqueo de cuenta y personajes.", FontTypeNames.FONTTYPE_USERBRONCE)
        
        If MessageNewbie <> vbNullString Then
            Call WriteConsoleMsg(UserIndex, MessageNewbie, FontTypeNames.FONTTYPE_CONSEJOVesA)
        End If
        
        If .GuildIndex > 0 Then
            Call Guilds_Connect(UserIndex)
                  
        End If

        If (.flags.Muerto = 0) Then
            .flags.SeguroResu = False
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)
        Else
            .flags.SeguroResu = True
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)
        End If
        
        Call WriteMultiMessage(UserIndex, eMessages.DragSafeOff)
            
        If Escriminal(UserIndex) Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
        End If
        
        If .Stats.Gld < 0 Then .Stats.Gld = 0

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FXWARP, 0))
    
        Call WriteLoggedMessage(UserIndex)
    
          ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
        Call IntervaloPermiteSerAtacado(UserIndex, True)
    
        Call MostrarNumUsers
        '
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
        
        If MapInfo(.Pos.Map).LvlMin > .Stats.Elv Then
            Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, False)
        End If
        
        If MapInfo(.Pos.Map).OnLoginGoTo.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡No puedes circular por este mapa en estos momentos. Te llevare a un sitio seguro!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).OnLoginGoTo.Map, MapInfo(.Pos.Map).OnLoginGoTo.X, MapInfo(.Pos.Map).OnLoginGoTo.Y, True, True)
        End If
        
        If Not CheckMap_Onlines(UserIndex, .Pos) Then
        
            If MapInfo(.Pos.Map).GoToOns.Map > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡¡¡No puedes circular por este mapa en estos momentos. Te llevare a la entrada del mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
                Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).GoToOns.Map, MapInfo(.Pos.Map).GoToOns.X, MapInfo(.Pos.Map).GoToOns.Y, True, True)
            Else
                Call EventWarpUser(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)
            End If
            
        End If
        
        ' Chequea si está en un mapa por horario y lo regresa a la ciudad principal
        If Not CheckMap_HourDay(UserIndex, .Pos) Then
            Call EventWarpUser(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)
        End If
        
        
        
        If .flags.Envenenado = 1 Then
            Call WriteUpdateEffect(UserIndex)
        End If
        
        
        Call WriteSendIntervals(UserIndex)
              
          Call UpdatePremium(UserIndex)
          Call WriteConsoleMsg(UserIndex, "Tipea /SHOP para ingresar a la Tienda Oficial de la comunidad, donde podrás comprar por DSP", FontTypeNames.FONTTYPE_INFOBOLD)
              
          Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageUpdateEvento(EsModoEvento))
          Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageUpdateMeditation(.MeditationUser, Meditation(.MeditationSelected)))
        

        If EventLast > 0 Then
               Call WriteConsoleMsg(UserIndex, "Eventos> Nuevos eventos en curso. Tipea /TORNEOS para saber más.", FontTypeNames.FONTTYPE_CRITICO)
        End If
        

        
        If Not NewChar Then
            If .Stats.Elv <= LimiteNewbie Then
                Call WriteQuestInfo(UserIndex, True, 0)
                Call WriteConsoleMsg(UserIndex, "Misiones> Accede al panel de misiones desde la tecla 'ESC' o bien escribiendo /MISIONES", FontTypeNames.FONTTYPE_CRITICO)
            End If
        End If
        
        
        ' # Comprueba la permanencia de skins especiales (Clanes)
        Call Skins_CheckGuild(UserIndex, True)
        
        
        FlushBuffer UserIndex

    End With
    
    Exit Sub
ErrHandler:
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ResetFacciones_Err
        '</EhHeader>

        '*************************************************
        'Author: Unknown
        '*************************************************
100     With UserList(UserIndex).Faction
    
102         .FragsCiu = 0
104         .FragsCri = 0
106         .FragsOther = 0
108         .ExFaction = 0
110         .Range = 0
112         .Status = 0
114         .StartDate = vbNullString
116         .StartElv = 0
        End With

        '<EhFooter>
        Exit Sub

ResetFacciones_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetFacciones " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

        '<EhHeader>
        On Error GoTo ResetContadores_Err

        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 10/07/2010
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '05/20/2007 Integer - Agregue todas las variables que faltaban.
        '10/07/2010: ZaMa - Agrego los counters que faltaban.
        '*************************************************
100     With UserList(UserIndex).Counters
            .SpeedHackCounter = 0
            .LastStep = 0
            .TimeGMBOT = 0
102         .controlHechizos.HechizosCasteados = 0
104         .controlHechizos.HechizosTotales = 0
106         .Incinerado = 0
108         .LastSave = 0
110         .TimeLastReset = 0
112         .PacketCount = 0
114         .RuidoPocion = 0
116         .RuidoDopa = 0
118         .SpamMessage = 0
120         .MessageSend = 0
122         .FightInvitation = 0
124         .FightSend = 0
126         .Drawers = 0
128         .DrawersCount = 0
130         .TimeInfoMao = 0
132         .TimeDrop = 0
              .TimeEquipped = 0
134         .TimerPuedeCastear = 0
            .TimerPuedeRecibirAtaqueCriature = 0
136         .TimeInfoChar = 0
138         .TimeCommerce = 0
140         .TimeMessage = 0
142         .TimeInfoMao = 0
144         .TimePublicationMao = 0
        
146         .TimeInactive = 0
148         .TimeBono = 0
150         .TimeTelep = 0
152         .TimeApparience = 0
154         .TimeFight = 0
156         .TimeCreateChar = 0
        
158         .AGUACounter = 0
160         .AsignedSkills = 0
162         .AttackCounter = 0
166         .Ceguera = 0
168         .COMCounter = 0
170         .Estupidez = 0
172         .failedUsageAttempts = 0
174         .failedUsageAttempts_Clic = 0
176         .failedUsageCastSpell = 0
178         .Frio = 0
180         .goHome = 0
182         .goHomeSec = 0
184         .HPCounter = 0
186         .IdleCount = 0
188         .Invisibilidad = 0
190         .Lava = 0
192         .Mimetismo = 0
194         .Ocultando = 0
196         .Paralisis = 0
198         .Pena = 0
200         .PiqueteC = 0
202         .Saliendo = False
204         .Salir = 0
206         .STACounter = 0
208         .TiempoOculto = 0
210         .TimerEstadoAtacable = 0
212         .TimerGolpeMagia = 0
214         .TimerGolpeUsar = 0
216         .TimerUsarClick = 0
218         .TimerLanzarSpell = 0
            .BuffoAceleration = 0
            .TimerShiftear = 0
            .CaspeoTime = 0
220         .TimerMagiaGolpe = 0
222         .TimerPerteneceNpc = 0
224         .TimerPuedeAtacar = 0
226         .TimerPuedeSerAtacado = 0
228         .TimerPuedeTrabajar = 0
230         .TimerPuedeUsarArco = 0
232         .TimerUsar = 0
234         .Trabajando = 0
236         .Veneno = 0

        End With

        '<EhFooter>
        Exit Sub

ResetContadores_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetContadores " & "at line " & Erl

        '</EhFooter>
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ResetCharInfo_Err
        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
100     With UserList(UserIndex).Char
102         .Body = 0
104         .CascoAnim = 0
106         .charindex = 0
108         .FX = 0
110         .Head = 0
112         .loops = 0
114         .Heading = 0
116         .loops = 0
118         .ShieldAnim = 0
120         .WeaponAnim = 0
              .speeding = 0
        End With

        '<EhFooter>
        Exit Sub

ResetCharInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetCharInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ResetBasicUserInfo_Err
        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
100     With UserList(UserIndex)
102         .Name = vbNullString
104         .secName = vbNullString
106         .Desc = vbNullString
108         .DescRM = vbNullString
110         .Pos.Map = 0
112         .Pos.X = 0
114         .Pos.Y = 0
116         .Clase = 0
118         .Genero = 0
120         .Hogar = 0
122         .Raza = 0
        
124         .GroupIndex = 0
126         .GroupRequest = vbNullString
128         .GroupRequestTime = 0
130         .GroupSlotUser = 0
        
132         With .Stats
134             .Elv = 0
136             .Elu = 0
138             .Exp = 0


140             .Armour = 0
                  .ArmourMag = 0
                  .Damage = 0
                  .DamageMag = 0
                  .RegHP = 0
                  .RegMANA = 0
                  .Cooldown = 0
                  .Attack = 0
                  .Movement = 0
                  
                  
                  
142             .NPCsMuertos = 0
146             .SkillPts = 0
148             .Gld = 0
150             .UserAtributos(1) = 0
152             .UserAtributos(2) = 0
154             .UserAtributos(3) = 0
156             .UserAtributos(4) = 0
158             .UserAtributos(5) = 0
160             .UserAtributosBackUP(1) = 0
162             .UserAtributosBackUP(2) = 0
164             .UserAtributosBackUP(3) = 0
166             .UserAtributosBackUP(4) = 0
168             .UserAtributosBackUP(5) = 0
            End With
        
        End With

        '<EhFooter>
        Exit Sub

ResetBasicUserInfo_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetBasicUserInfo " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ResetReputacion_Err
        '</EhHeader>

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
100     With UserList(UserIndex).Reputacion
102         .AsesinoRep = 0
104         .BandidoRep = 0
106         .BurguesRep = 0
108         .LadronesRep = 0
110         .NobleRep = 0
112         .PlebeRep = 0
114         .NobleRep = 0
116         .promedio = 0
        End With

        '<EhFooter>
        Exit Sub

ResetReputacion_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetReputacion " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetUserMeditation(ByRef IUser As User)
        '<EhHeader>
        On Error GoTo ResetUserMeditation_Err
        '</EhHeader>

        Dim A As Long
    
100     With IUser

102         For A = 1 To MAX_MEDITATION
104             .MeditationUser(A) = 0
106         Next A
        
108         .MeditationSelected = 0
        End With

        '<EhFooter>
        Exit Sub

ResetUserMeditation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetUserMeditation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetUserOld(ByRef IUser As User)
        '<EhHeader>
        On Error GoTo ResetUserOld_Err
        '</EhHeader>
    
100     With IUser.OldInfo
102         .Clase = 0
104         .Raza = 0
106         .GldBlue = 0
108         .GldRed = 0
110         .MaxHp = 0
112         .MaxMan = 0
              .MaxSta = 0
              .Elv = 0
              .Exp = 0
              .Head = 0
              
            Dim A As Long

114         For A = 1 To MAXUSERHECHIZOS
116             .UserSpell(A) = 0
118         Next A
    
        End With
    
        '<EhFooter>
        Exit Sub

ResetUserOld_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetUserOld " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetUserObjectClaim(ByRef IUser As User)
        '<EhHeader>
        On Error GoTo ResetUserObjectClaim_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To MAX_INVENTORY_SLOTS
        
102         With IUser.ObjectClaim(A)
104             .Amount = 0
106             .ObjIndex = 0
108             .Equipped = 0
            End With
        
110     Next A

        '<EhFooter>
        Exit Sub

ResetUserObjectClaim_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetUserObjectClaim " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetCharacterStats(ByRef IUser As User)

    With IUser.CharacterStats
        .PassiveAccumulated = 0
        
    End With
    
End Sub
Sub ResetUserFlags(ByRef IUser As User)

        '*************************************************
        'Author: Unknown
        'Last modified: 06/28/2008
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
        '*************************************************
        '<EhHeader>
        On Error GoTo ResetUserFlags_Err

        '</EhHeader>
            
        Dim A As Long
        
        IUser.QuestLast = 0
        ReDim IUser.QuestStats(1 To MAXUSERQUESTS) As tUserQuest
          
100     Call ResetUserMeditation(IUser)
102     Call ResetUserOld(IUser)
104     Call Reto_ResetUserTemp(IUser)
106     Call ResetUserObjectClaim(IUser)
        Call AntiFrags_ResetInfo(IUser)
        
        Call ResetCharacterStats(IUser)
          
        ' # Reset quests
        For A = 1 To MAXUSERQUESTS
108        Call CleanQuestSlot(IUser, A)
        Next A
          
        Dim NullBot As tBotIntelligence
            
        For A = 1 To BOT_MAX_USER
            
            IUser.BotIntelligence(A) = NullBot
        Next A

        Dim i As Long
            
110     With IUser
            
            .PosOculto.Map = 0
            .PosOculto.X = 0
            .PosOculto.Y = 0
            
            ReDim .Skins.ObjIndex(1 To MAX_INVENTORY_SKINS) As Integer
            
            .Skins.Last = 0
            .Skins.ArmourIndex = 0
            .Skins.HelmIndex = 0
            .Skins.ShieldIndex = 0
            .Skins.WeaponIndex = 0
            .Skins.WeaponArcoIndex = 0
            .Skins.WeaponDagaIndex = 0
            
            For i = 1 To MAX_INVENTORY_SKINS
                .Skins.ObjIndex(i) = 0
            Next i
            
            .GuildIndex = 0
            .GuildRange = 0
            .GuildSlot = 0
118         .UseObj_Clic = 0
120         .UseObj_Init_Clic = 0
122         .UseObj_U = 0
124         .UseObj_Init_U = 0
126         .Next_UseItem = False
        
128         .LastPotion = eLastPotion.eNullPotion
130         .PotionBlue_Clic = 0
132         .PotionBlue_Clic_Interval = 0
134         .PotionBlue_U = 0
136         .PotionBlue_U_Interval = 0
138         .PotionRed_Clic = 0
140         .PotionRed_Clic_Interval = 0
142         .PotionRed_U = 0
144         .PotionRed_U_Interval = 0
        
146         For A = 0 To 1
148             .interval(A).IAttack = 0
150             .interval(A).IDrop = 0
152             .interval(A).ISpell = 0
154             .interval(A).IUse = 0
156             .interval(A).ILeftClick = 0
158         Next A
        
160         .MascotaIndex = 0
162         .DañoApu = 0
164         .UserKey = 0
166         .Power = False
168         .UserLastClick = 0
170         .UserLastClick_Tolerance = 0
        
172         With .Stats
174             .Points = 0

            End With

        End With
    
176     With IUser.flags

            .RachasTemp = 0
            .Rachas = 0
            .RedLimit = 0
            .RedUsage = 0
            .RedValid = False
            .BotList = 0
            .TeleportInvoker = 0
            .LastInvoker = 0
            .TempAccount = vbNullString
            .TempPasswd = vbNullString
            .DeslogeandoCuenta = False
            .StreamUrl = vbNullString
            .ModoStream = False
178         .ToleranceCheat = 0
180         .DragBlocked = False
182         .GmSeguidor = 0
184         .LastSlotClient = 0
186         .MenuCliente = 0
188         .Montando = 0
190         .DesafiosGanados = 0
192         .Desafiando = 0
194         .SelectedBono = 0
196         .Premium = 0
198         .Streamer = 0
200         .Bronce = 0
202         .Transform = 0
204         .TransformVIP = 0
206         .Plata = 0
208         .Oro = 0
210         .SlotReto = 0
212         .SlotEvent = 0
214         .SlotUserEvent = 0
216         .SlotRetoUser = 255
218         .SlotFast = 0
220         .SlotFastUser = 0
222         .SelectedEvent = 0
224         .FightTeam = 0
226         .Comerciando = False
228         .Ban = 0
230         .Escondido = 0
232         .DuracionEfecto = 0
234         .NpcInv = 0
236         .StatsChanged = 0
238         .TargetNPC = 0
240         .TargetNpcTipo = eNPCType.Comun
242         .TargetObj = 0
244         .TargetObjMap = 0
246         .TargetObjX = 0
248         .TargetObjY = 0
250         .TargetUser = 0
252         .TipoPocion = 0
254         .TomoPocion = False
256         .Hambre = 0
258         .Sed = 0
260         .Descansar = False
262         .Vuela = 0
264         .Navegando = 0
266         .Oculto = 0
268         .Envenenado = 0
270         .Invisible = 0
272         .Paralizado = 0
274         .Inmovilizado = 0
276         .Maldicion = 0
278         .Bendicion = 0
280         .Meditando = 0
282         .Privilegios = 0
284         .PrivEspecial = False
286         .PuedeMoverse = 0
288         .OldBody = 0
290         .OldHead = 0
292         .AdminInvisible = 0
294         .ValCoDe = 0
296         .Hechizo = 0
304         .Silenciado = 0
306         .AdminPerseguible = False
308         .LastMap = 0
310         .Traveling = 0
312         .AtacablePor = 0
314         .AtacadoPorNpc = 0
316         .AtacadoPorUser = 0
318         .NoPuedeSerAtacado = False
320         .ShareNpcWith = 0
322         .EnConsulta = False
324         .Ignorado = False
326         .SendDenounces = False
328         .ParalizedBy = vbNullString
330         .ParalizedByIndex = 0
332         .ParalizedByNpcIndex = 0
        
        End With

        '<EhFooter>
        Exit Sub

ResetUserFlags_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserFlags " & "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetUserSpells_Err
        '</EhHeader>

        Dim LoopC As Long

100     For LoopC = 1 To MAXUSERHECHIZOS
102         UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
104     Next LoopC

        '<EhFooter>
        Exit Sub

ResetUserSpells_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetUserSpells " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetUserBanco_Err
        '</EhHeader>

        Dim LoopC As Long
    
100     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
102         UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
104         UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
106         UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
108     Next LoopC
    
110     UserList(UserIndex).BancoInvent.NroItems = 0
        '<EhFooter>
        Exit Sub

ResetUserBanco_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetUserBanco " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo LimpiarComercioSeguro_Err
        '</EhHeader>

100     With UserList(UserIndex).ComUsu

102         If .DestUsu > 0 Then
104             Call FinComerciarUsu(.DestUsu)
106             Call FinComerciarUsu(UserIndex)
            End If

        End With

        '<EhFooter>
        Exit Sub

LimpiarComercioSeguro_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.LimpiarComercioSeguro " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetUserSlot_Err
        '</EhHeader>

        Dim i As Long

100     Call LimpiarComercioSeguro(UserIndex)
102     Call ResetFacciones(UserIndex)
104     Call ResetContadores(UserIndex)
105     Call ResetPacketRateData(UserIndex)
106     Call ResetCharInfo(UserIndex)
108     Call ResetBasicUserInfo(UserIndex)
110     Call ResetReputacion(UserIndex)
112     Call ResetUserFlags(UserList(UserIndex))
114     Call ResetKeyPackets(UserIndex)
116     Call ResetPointer(UserIndex, Point_Inv)
118     Call ResetPointer(UserIndex, Point_Spell)
120     Call LimpiarInventario(UserIndex)
122     Call ResetUserSpells(UserIndex)
124     Call ResetUserBanco(UserIndex)
126     Call LimpiarMensajes(UserIndex)

128     With UserList(UserIndex).ComUsu
130         .Acepto = False
    
132         For i = 1 To MAX_OFFER_SLOTS
134             .cant(i) = 0
136             .Objeto(i) = 0
138         Next i
        
140         .EldhirAmount = 0
142         .GoldAmount = 0
144         .DestNick = vbNullString
146         .DestUsu = 0
        End With
        
334         If UserList(UserIndex).flags.OwnedNpc <> 0 Then
336             Call PerdioNpc(UserIndex)
            End If
 
        '<EhFooter>
        Exit Sub

ResetUserSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.ResetUserSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub CloseUser(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CloseUser_Err
        '</EhHeader>


        Dim N    As Integer

        Dim Map  As Integer

        Dim Name As String

        Dim i    As Integer

        Dim aN   As Integer

100     With UserList(UserIndex)
102         aN = .flags.AtacadoPorNpc

104         If aN > 0 Then
106             Npclist(aN).Movement = Npclist(aN).flags.OldMovement
108             Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
110             Npclist(aN).flags.AttackedBy = vbNullString
                  Npclist(aN).Target = 0
            End If
    
112         aN = .flags.NPCAtacado

114         If aN > 0 Then
116             If Npclist(aN).flags.AttackedFirstBy = .Name Then
118                 Npclist(aN).flags.AttackedFirstBy = vbNullString
                End If
            End If

120         .flags.AtacadoPorNpc = 0
122         .flags.NPCAtacado = 0
    
124         Map = .Pos.Map
126         Name = UCase$(.Name)
    
128         .Char.FX = 0
130         .Char.loops = 0
132         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 0, 0))
    
134         .flags.UserLogged = False
136         .Counters.Saliendo = False
    
            'Le devolvemos el body y head originales
138         If .flags.AdminInvisible = 1 Then
140             .Char.Body = .flags.OldBody
142             .Char.Head = .flags.OldHead
144             .flags.AdminInvisible = 0
            End If
    
            'si esta en party le devolvemos la experiencia
146         If .GroupIndex > 0 Then Call mGroup.AbandonateGroup(UserIndex)
    
            'Save statistics
            'Call Statistics.UserDisconnected(UserIndex)
    
            ' Grabamos el personaje del usuario
148         Call SaveUser(UserList(UserIndex), CharPath & Name & ".chr")

            'Quitar el dialogo
            'If MapInfo(Map).NumUsers > 0 Then
            '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
            'End If
    
150             If MapInfo(Map).NumUsers > 0 Then
152                 Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
                End If
        
                'Update Map Users
154             MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
156             MapInfo(Map).Players.Remove UserIndex
    
158             If MapInfo(Map).NumUsers < 0 Then
160                 MapInfo(Map).NumUsers = 0
                End If
            'End If
    
            'Borrar el personaje
162         If .Char.charindex > 0 Then
164             Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
            End If
    
            'Borrar mascota
166         If .MascotaIndex Then
168             If Npclist(.MascotaIndex).flags.NPCActive Then Call QuitarNPC(.MascotaIndex)
            End If
            
            
            ' Remove Position
            Call Guilds_UpdatePosition(UserIndex)

        End With



        '<EhFooter>
        Exit Sub

CloseUser_Err:
        LogError Err.description & vbCrLf & _
               "in CloseUser " & _
               "at line " & Erl

        '</EhFooter>
End Sub

Sub ReloadSokcet()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    'Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    'If NumUsers <= 0 Then
    'Call WSApiReiniciarSockets
    'Else
    '       Call apiclosesocket(SockListen)
    '       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    'End If

    Exit Sub

ErrHandler:
    Call LogError("Error en CheckSocketState " & Err.number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo EcharPjsNoPrivilegiados_Err
        '</EhHeader>

        Dim LoopC As Long
    
100     For LoopC = 1 To LastUser

102         If UserList(LoopC).flags.UserLogged Then
104             If UserList(LoopC).flags.Privilegios And PlayerType.User Then
106                 Call Protocol.Kick(LoopC)
            
                End If
            End If

108     Next LoopC

        '<EhFooter>
        Exit Sub

EcharPjsNoPrivilegiados_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.TCP.EcharPjsNoPrivilegiados " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub User_GenerateNewHead(ByVal UserIndex As Integer, ByVal Tipe As Byte)

        '<EhHeader>
        On Error GoTo User_GenerateNewHead_Err

        '</EhHeader>
    
        Dim NewHead    As Integer

        Dim UserRaza   As Byte

        Dim UserGenero As Byte
    
100     UserGenero = UserList(UserIndex).Genero
102     UserRaza = UserList(UserIndex).Raza

104     Select Case UserGenero

            Case eGenero.Hombre

106             Select Case UserRaza

                    Case eRaza.Humano
                        
                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(502, 546)
                        Else
                            NewHead = RandomNumber(1, 25)

                        End If

108
                    
110                 Case eRaza.Elfo

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(577, 608)
                        Else
                            NewHead = RandomNumber(102, 111)

                        End If
                        
114                 Case eRaza.Drow
                        '
116

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(639, 669)
                        Else
                            NewHead = RandomNumber(201, 205)

                        End If
                        
118                 Case eRaza.Enano

                        '
                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(700, 729)
                        Else
                            NewHead = RandomNumber(301, 305)

                        End If

122                 Case eRaza.Gnomo

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(760, 789)
                        Else
                            NewHead = RandomNumber(401, 405)

                        End If

124

                End Select

126         Case eGenero.Mujer

128             Select Case UserRaza

                    Case eRaza.Humano

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(547, 576)
                        Else
                            NewHead = RandomNumber(71, 75)

                        End If
                    
132                 Case eRaza.Elfo

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(609, 638)
                        Else
                            NewHead = RandomNumber(170, 176)

                        End If

136                 Case eRaza.Drow

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(670, 699)
                        Else
                            NewHead = RandomNumber(270, 276)

                        End If

140                 Case eRaza.Gnomo

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(790, 819)
                        Else
                            NewHead = RandomNumber(471, 475)

                        End If

144                 Case eRaza.Enano
                        '
146

                        If Tipe = eEffectObj.e_NewHead Then
                            NewHead = RandomNumber(730, 759)
                        Else
                            NewHead = RandomNumber(370, 371)

                        End If

                End Select

        End Select
    
148     UserList(UserIndex).Char.Head = NewHead
150     UserList(UserIndex).OrigChar.Head = NewHead
    
        '<EhFooter>
        Exit Sub

User_GenerateNewHead_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.User_GenerateNewHead " & "at line " & Erl

        

        '</EhFooter>
End Sub

Sub ResetPacketRateData(ByVal UserIndex As Integer)

        On Error GoTo ResetPacketRateData_Err

        Dim i As Long
        
        With UserList(UserIndex)
        
            For i = 1 To MAX_PACKET_COUNTERS
                .MacroIterations(i) = 0
                .PacketTimers(i) = 0
                .PacketCounters(i) = 0
            Next i
            
        End With
        
        Exit Sub
        
ResetPacketRateData_Err:
End Sub
