Attribute VB_Name = "Mod_Balance"
Option Explicit

Private Const MAX_NIVEL_VIDA_PROMEDIO = 17

Public Type tRango

    minimo As Integer
    maximo As Integer

End Type

Private Type tBalance

    Hp As Integer
    Man As Integer

End Type

Public BalanceStats(1 To NUMCLASES, 1 To NUMRAZAS) As tBalance

' Adicionales de Vida
Public Const AdicionalHPGuerrero = 2 'HP adicionales cuando sube de nivel

Public Const AdicionalHPCazador = 1

Public Const AdicionalSTLadron = 3

Public Const AdicionalSTLeñador = 23

Public Const AdicionalSTPescador = 20

Public Const AdicionalSTMinero = 25

Public Const AumentoSTDef        As Byte = 15

Public Const AumentoStBandido    As Byte = AumentoSTDef + 23

Public Const AumentoSTLadron     As Byte = AumentoSTDef + 3

Public Const AumentoSTMago       As Byte = AumentoSTDef - 1

Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25

' El balance LVL TEMP sería el nivel "1" del personaje. En base al balance default del juego (TDS)
' Si se suma 32 + los 15 niveles que el juego tiene da 47. No es un número al azar
Public Const BALANCE_LVL_TEMP    As Byte = 0
Public Const POCION_ROJA_NEWBIE = 205
Public Const POCION_AZUL_NEWBIE = 206
Public Const POCION_AMARILLA_NEWBIE = 207
Public Const POCION_VERDE_NEWBIE = 208
Public Const VESTIMENTA_WAR_NEWBIE = 203
Public Const VESTIMENTA_MAG_NEWBIE = 204
Public Const DAGA_NEWBIE = 202

Public Sub LoadSetInitial_Class(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo LoadSetInitial_Class_Err
        '</EhHeader>

        Dim UserClase As eClass

        Dim Slot      As Long

100     With UserList(UserIndex)

102         Call LimpiarInventario(UserIndex)

            'Pociones Rojas (Newbie)
104         Slot = 1
106         .Invent.Object(Slot).ObjIndex = POCION_ROJA_NEWBIE
108         .Invent.Object(Slot).Amount = 150
        
            'Pociones azules (Newbie)
110         If .Stats.MaxMan > 0 Then
112             Slot = Slot + 1
114             .Invent.Object(Slot).ObjIndex = POCION_AZUL_NEWBIE
116             .Invent.Object(Slot).Amount = 100
          
            End If
        
            
122         .Invent.Object(Slot).Amount = 10 'Pociones Amarillas
118         Slot = Slot + 1
120         .Invent.Object(Slot).ObjIndex = POCION_AMARILLA_NEWBIE
        
            'Pociones Amarillas
124         Slot = Slot + 1
126         .Invent.Object(Slot).ObjIndex = POCION_VERDE_NEWBIE
128         .Invent.Object(Slot).Amount = 10
    
            Dim Escudo     As Obj

            Dim Casco      As Obj

            Dim Armadura   As Obj

            Dim Arma       As Obj

            Dim Anillo     As Obj

            Dim Municiones As Obj
    
130         Escudo.Amount = 1
132         Casco.Amount = 1
134         Armadura.Amount = 1
136         Arma.Amount = 1
138         Anillo.Amount = 1
140         Municiones.Amount = 1
             
142         .Char.WeaponAnim = NingunArma
144         .Char.CascoAnim = NingunCasco
146         .Char.Body = 0
148         .Char.ShieldAnim = NingunEscudo
         
150         Select Case .Clase
    
                Case eClass.Mage
152                 Escudo.ObjIndex = 0
154                 Casco.ObjIndex = 173 ' Sombrero de Mago
156                 Armadura.ObjIndex = 91 'Tunica legendaria
158                 Arma.ObjIndex = 171 ' Báculo Engarzado
            
160             Case eClass.Cleric
162                 Escudo.ObjIndex = 117 ' Escudo Imperial
164                 Casco.ObjIndex = 119 ' Casco de Hierro Completo
166                 Armadura.ObjIndex = 104 ' Placas de Acero
168                 Arma.ObjIndex = 141 ' Hacha Dos Filos
        
170             Case eClass.Paladin
172                 Escudo.ObjIndex = 117 ' Escudo Imperial
174                 Casco.ObjIndex = 119 ' Casco de Hierro Completo
176                 Armadura.ObjIndex = 104 ' Placas de Acero
178                 Arma.ObjIndex = 142 ' Espada de Plata
        
180             Case eClass.Warrior
182                 Escudo.ObjIndex = 117 ' Escudo Imperial
184                 Casco.ObjIndex = 119 ' Casco de Hierro Completo
186                 Armadura.ObjIndex = 104 ' Placas de Acero
188                 Arma.ObjIndex = 142 ' Espada de Plata
        
190             Case eClass.Assasin
192                 Escudo.ObjIndex = 333 ' Escudo de Tortuga
194                 Casco.ObjIndex = 118 ' Casco de Hierro
196                 Armadura.ObjIndex = 105 ' Armadura Klox
198                 Arma.ObjIndex = 139 ' Puñal Infernal
            
200             Case eClass.Bard
202                 Escudo.ObjIndex = 333 ' Escudo de Tortuga
204                 Casco.ObjIndex = 132 ' Casco de Hierro
206                 Armadura.ObjIndex = 91 'Tunica legendaria
208                 Arma.ObjIndex = 0
210                 Anillo.ObjIndex = 167 ' Laud Mágico
            
212             Case eClass.Druid
214                 Escudo.ObjIndex = 0
216                 Casco.ObjIndex = 0
218                 Armadura.ObjIndex = 91 'Tunica legendaria
220                 Arma.ObjIndex = 0
222                 Anillo.ObjIndex = 168 ' Anillo Mágico
            
224             Case eClass.Hunter
226                 Escudo.ObjIndex = 333 ' Escudo de Tortuga
228                 Casco.ObjIndex = 131 ' Capucha de Cazador
230                 Armadura.ObjIndex = 96 ' Armadura de Cazador
232                 Arma.ObjIndex = 212 ' Arco de Cazador
234                 Municiones.ObjIndex = 154 ' Flechas
        
236             Case Else
            
            End Select

238         If UserList(UserIndex).ServerSelected = 3 Then
240             Armadura.ObjIndex = 216 ' Vestimentas Polar

            End If
            
242         If Armadura.ObjIndex > 0 Then
244             Slot = Slot + 1
246             .Invent.Object(Slot).ObjIndex = Armadura.ObjIndex
    
248             .Invent.Object(Slot).Amount = 1
250             .Invent.Object(Slot).Equipped = 1
          
252             .Invent.ArmourEqpSlot = Slot
254             .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
256             .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
        
            End If
    
258         If Arma.ObjIndex > 0 Then
260             Slot = Slot + 1

262             .Invent.Object(Slot).ObjIndex = Arma.ObjIndex

264             .Invent.Object(Slot).Amount = 1
266             .Invent.Object(Slot).Equipped = 1
          
268             .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
270             .Invent.WeaponEqpSlot = Slot
          
272             .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, .Invent.WeaponEqpObjIndex)
    
            End If
    
274         If Municiones.ObjIndex > 0 Then
276             Slot = Slot + 1

278             .Invent.Object(Slot).ObjIndex = Municiones.ObjIndex

280             .Invent.Object(Slot).Amount = 1
282             .Invent.Object(Slot).Equipped = 1
          
284             .Invent.MunicionEqpObjIndex = .Invent.Object(Slot).ObjIndex
286             .Invent.MunicionEqpSlot = Slot
    
            End If
            
288         If Escudo.ObjIndex > 0 Then
290             Slot = Slot + 1

292             .Invent.Object(Slot).ObjIndex = Escudo.ObjIndex

294             .Invent.Object(Slot).Amount = 1
296             .Invent.Object(Slot).Equipped = 1
          
298             .Invent.EscudoEqpObjIndex = .Invent.Object(Slot).ObjIndex
300             .Invent.EscudoEqpSlot = Slot
          
302             .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

            End If
    
304         If Casco.ObjIndex > 0 Then
306             Slot = Slot + 1

308             .Invent.Object(Slot).ObjIndex = Casco.ObjIndex

310             .Invent.Object(Slot).Amount = 1
312             .Invent.Object(Slot).Equipped = 1
          
314             .Invent.CascoEqpObjIndex = .Invent.Object(Slot).ObjIndex
316             .Invent.CascoEqpSlot = Slot
          
318             .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
    
            End If
        
            ' Total Items
320         .Invent.NroItems = Slot
        
322         Call User_GenerateNewHead(UserIndex, 1)
324         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
        
326         Call UpdateUserInv(True, UserIndex, 0)

            ' Spells
            Dim A As Long

328         For A = 1 To MAXUSERHECHIZOS
330             .Stats.UserHechizos(A) = 0
332         Next A
            
334         If .Stats.MaxMan > 0 Then
336             .Stats.UserHechizos(35) = 10 'RemoverParalisis
338             .Stats.UserHechizos(34) = 24 ' 'Inmovilizar
340             .Stats.UserHechizos(33) = 9 ' 'Paralizar
342             .Stats.UserHechizos(32) = 15 ' 'Tormenta de Fuego
344             .Stats.UserHechizos(31) = 23 ' 'Descarga Eléctrica
346             .Stats.UserHechizos(30) = 14 ' 'Invisibilidad
                
348             If .Stats.MaxMan > 1000 Then
350                 .Stats.UserHechizos(29) = 25 ' 'Apocalipsis

                End If

            End If

352         Call UpdateUserHechizos(True, UserIndex, 0)
        
        End With

        '<EhFooter>
        Exit Sub

LoadSetInitial_Class_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.LoadSetInitial_Class " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ApplySetInitial_Newbie(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ApplySetInitial_Newbie_Err
        '</EhHeader>

100     With UserList(UserIndex)

            '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
            Dim Slot      As Byte

            Dim IsPaladin As Boolean
        
102         IsPaladin = .Clase = eClass.Paladin
        
            'Pociones Rojas (Newbie)
104         Slot = 1
106         .Invent.Object(Slot).ObjIndex = POCION_ROJA_NEWBIE
108         .Invent.Object(Slot).Amount = 150
        
110         If POCION_AZUL_NEWBIE > 0 Then

                'Pociones azules (Newbie)
112             If .Stats.MaxMan > 0 Or IsPaladin Then
114                 Slot = Slot + 1
116                 .Invent.Object(Slot).ObjIndex = POCION_AZUL_NEWBIE
118                 .Invent.Object(Slot).Amount = 100
              
                End If

            End If
        
120         If POCION_AMARILLA_NEWBIE > 0 Then
                'Pociones Amarillas
122             Slot = Slot + 1
124             .Invent.Object(Slot).ObjIndex = POCION_AMARILLA_NEWBIE
126             .Invent.Object(Slot).Amount = 10

            End If
        
128         If POCION_VERDE_NEWBIE > 0 Then
                'Pociones Amarillas
130             Slot = Slot + 1
132             .Invent.Object(Slot).ObjIndex = POCION_VERDE_NEWBIE
134             .Invent.Object(Slot).Amount = 10

            End If
        
            ' Ropa (Newbie)
136         Slot = Slot + 1
        

138             Select Case .Clase

                    Case eClass.Assasin, eClass.Paladin, eClass.Hunter, eClass.Warrior, eClass.Thief
140                     .Invent.Object(Slot).ObjIndex = VESTIMENTA_WAR_NEWBIE

142                 Case eClass.Mage, eClass.Druid, eClass.Cleric, eClass.Bard
144                     .Invent.Object(Slot).ObjIndex = VESTIMENTA_MAG_NEWBIE

                End Select

        
            ' Equipo ropa
156         .Invent.Object(Slot).Amount = 1
158         .Invent.Object(Slot).Equipped = 1
160         .flags.Desnudo = 0
162         .Invent.ArmourEqpSlot = Slot
164         .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
            'Call DarCuerpoDesnudo(UserIndex)
166         .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
        
168         If DAGA_NEWBIE > 0 Then
                'Arma (Newbie)
170             Slot = Slot + 1
    
172             .Invent.Object(Slot).ObjIndex = DAGA_NEWBIE
            
                ' Equipo arma
174             .Invent.Object(Slot).Amount = 1
176             .Invent.Object(Slot).Equipped = 1
        
178             .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
180             .Invent.WeaponEqpSlot = Slot
            
182             .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, .Invent.WeaponEqpObjIndex)
            End If

            ' Sin casco y escudo
184         .Char.ShieldAnim = NingunEscudo
186         .Char.CascoAnim = NingunCasco
          
            ' Total Items
188         .Invent.NroItems = Slot

        End With

        '<EhFooter>
        Exit Sub

ApplySetInitial_Newbie_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.ApplySetInitial_Newbie " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ApplySpellsStats(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ApplySpellsStats_Err
        '</EhHeader>

        ' Aplicamos hechizos
100     With UserList(UserIndex)

102         If .Clase = eClass.Mage Or _
                .Clase = eClass.Cleric Or _
                .Clase = eClass.Druid Or _
                .Clase = eClass.Bard Or _
                .Clase = eClass.Assasin Or _
                .Clase = eClass.Paladin Then
            
104             .Stats.UserHechizos(1) = 2      ' Dardo Mágico
                .Stats.UserHechizos(2) = 1      ' Curar Veneno
                .Stats.UserHechizos(3) = 3      ' Curar heridas Leves
            End If

        End With

        '<EhFooter>
        Exit Sub

ApplySpellsStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.ApplySpellsStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub InitialUserStats(ByRef IUser As User)
        '<EhHeader>
        On Error GoTo InitialUserStats_Err
        '</EhHeader>
    
        ' Reset automático para incrementar de forma automática
        ' Adaptado a 0.11.5
        Dim LoopC As Integer

        Dim MiInt As Long

        Dim ln    As String
            
100     With IUser
    
102         .Stats.UserAtributos(eAtributos.Fuerza) = 18 + Balance.ModRaza(.Raza).Fuerza
104         .Stats.UserAtributos(eAtributos.Agilidad) = 18 + Balance.ModRaza(.Raza).Agilidad
106         .Stats.UserAtributos(eAtributos.Inteligencia) = 18 + Balance.ModRaza(.Raza).Inteligencia
108         .Stats.UserAtributos(eAtributos.Carisma) = 18 + Balance.ModRaza(.Raza).Carisma
110         .Stats.UserAtributos(eAtributos.Constitucion) = 18 + Balance.ModRaza(.Raza).Constitucion
        
            ' Skills en 0
112         For LoopC = 1 To NUMSKILLS
114             .Stats.UserSkills(LoopC) = 0
116             'Call CheckEluSkill(UserIndex, LoopC, True)
118         Next LoopC

            ' Vida
            'MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
120         .Stats.MaxHp = 20 ' + MiInt
122         .Stats.MinHp = 20 ' + MiInt

            ' Energia
124         MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)

126         If MiInt = 1 Then MiInt = 2
128         .Stats.MaxSta = 20 * MiInt
130         .Stats.MinSta = 20 * MiInt
          
            ' Agua y comida
132         .Stats.MaxAGU = 100
134         .Stats.MinAGU = 100
136         .Stats.MaxHam = 100
138         .Stats.MinHam = 100

            '<-----------------MANA----------------------->
140         If .Clase = eClass.Mage Then  'Cambio en mana inicial (ToxicWaste)
142             MiInt = 100
144             .Stats.MaxMan = MiInt
146             .Stats.MinMan = MiInt
148         ElseIf .Clase = eClass.Cleric Or .Clase = eClass.Druid Or .Clase = eClass.Bard Or .Clase = eClass.Assasin Then
150             .Stats.MaxMan = 50
152             .Stats.MinMan = 50
            Else
154             .Stats.MaxMan = 0
156             .Stats.MinMan = 0

            End If

158         .Stats.MinMan = .Stats.MaxMan
        
160         .Stats.MaxHit = 2
162         .Stats.MinHit = 1

164         .Stats.Exp = 0
166         .Stats.Elu = 300
168         .Stats.Elv = 1
170         .Stats.SkillPts = 10
        End With

        '<EhFooter>
        Exit Sub

InitialUserStats_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.InitialUserStats " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub UserLevelEditation(ByRef IUser As User, ByVal Elv As Byte, ByVal UserUps As Byte)
    ' Procedimiento creado para entrenamiento de personajes de nivel 1 a 15. (Editar con f1)
    
    Dim LoopC  As Integer

    Dim NewHp  As Integer

    Dim NewMan As Integer

    Dim NewSta As Integer

    Dim NewHit As Integer
    
    On Error GoTo UserLevelEditation_Error

    With IUser
        
        ' Quitamos los Itmes Newbies
        'Call QuitarNewbieObj(UserIndex)
        
        NewMan = .Stats.MaxMan
        NewHp = 0
        NewHit = 0
        NewSta = 0
        
        'Nivel 2 a 15
        For LoopC = 2 To Elv
            NewMan = NewMan + Balance_AumentoMANA(.Clase, .Raza, NewMan)
            NewSta = NewSta + Balance_AumentoSTA(.Clase)
            NewHit = NewHit + Balance_AumentoHIT(.Clase, LoopC)
        Next LoopC
        
        
        ' Nueva vida
        .Stats.MaxHp = getVidaIdeal(Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion)) + UserUps
        
        If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
        
        ' Nueva energía
        .Stats.MaxSta = .Stats.MaxSta + NewSta

        If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

        ' Nueva maná
        .Stats.MaxMan = NewMan

        If .Stats.MaxMan > STAT_MAXMAN Then .Stats.MaxMan = STAT_MAXMAN
        
        ' Nuevo golpe máximo y mínimo
        .Stats.MaxHit = .Stats.MaxHit + NewHit
        .Stats.MinHit = .Stats.MinHit + NewHit

        If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then .Stats.MaxHit = STAT_MAXHIT_UNDER36
        If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then .Stats.MinHit = STAT_MAXHIT_UNDER36
        
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinSta = .Stats.MaxSta
        .Stats.Elv = Elv
        .Stats.Elu = EluUser(.Stats.Elv)
        .Stats.Gld = 0
        .Stats.MaxAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinAGU = .Stats.MaxAGU
        .Stats.MinHam = .Stats.MinHam
    End With

    On Error GoTo 0

    Exit Sub

UserLevelEditation_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure UserLevelEditation of Módulo mBalance in line " & Erl

End Sub

Public Function Balance_AumentoSTA(ByVal UserClase As Byte) As Integer
    ' Aumento de energía
    
    On Error GoTo Balance_AumentoSTA_Error
            
        Select Case UserClase

            Case eClass.Thief
                Balance_AumentoSTA = AumentoSTLadron

            Case eClass.Mage
                Balance_AumentoSTA = AumentoSTMago


            Case Else
                Balance_AumentoSTA = AumentoSTDef

        End Select
        
    On Error GoTo 0

    Exit Function

Balance_AumentoSTA_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure Balance_AumentoSTA of Módulo mBalance in line " & Erl

End Function

Public Function Balance_AumentoHIT(ByVal UserClase As Byte, _
                                   ByVal Elv As Byte) As Integer

    ' Aumento de HIT por nivel
    On Error GoTo Balance_AumentoHIT_Error

            
        Select Case UserClase

            Case eClass.Warrior, eClass.Hunter
                Balance_AumentoHIT = IIf(Elv > 35, 2, 3)
                    
            Case eClass.Paladin
                Balance_AumentoHIT = IIf(Elv > 35, 1, 3)
                    
            Case eClass.Thief
                Balance_AumentoHIT = 2
                         
            Case eClass.Mage
                Balance_AumentoHIT = 1
                   
            Case eClass.Cleric
                Balance_AumentoHIT = 2
                    
            Case eClass.Druid
                Balance_AumentoHIT = 2
                     
            Case eClass.Assasin
                Balance_AumentoHIT = IIf(Elv > 35, 1, 3)

            Case eClass.Bard
                Balance_AumentoHIT = 2
                    
            Case Else
                Balance_AumentoHIT = 2

        End Select
    On Error GoTo 0

    Exit Function

Balance_AumentoHIT_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure Balance_AumentoHIT of Módulo mBalance in line " & Erl

End Function

Public Function Balance_AumentoMANA(ByVal Class As Byte, ByVal Raze As Byte, ByRef TempMan As Integer) As Integer
        ' Aumento de maná según clase
        '<EhHeader>
        On Error GoTo Balance_AumentoMANA_Err
        '</EhHeader>
    
        Dim UserInteligencia As Byte

100     UserInteligencia = 18 + Balance.ModRaza(Raze).Inteligencia
    
        On Error GoTo Balance_AumentoMANA_Error

102     Select Case Class
                    
            Case eClass.Paladin
104             Balance_AumentoMANA = UserInteligencia
                         
106         Case eClass.Mage

108             If Raze = Enano Then
110                 Balance_AumentoMANA = 2 * UserInteligencia
112             ElseIf (TempMan >= 2000) Then
114                 Balance_AumentoMANA = (3 * UserInteligencia) / 2
                Else
116                 Balance_AumentoMANA = 3 * UserInteligencia

                End If
                   
118         Case eClass.Druid, eClass.Bard, eClass.Cleric
120             Balance_AumentoMANA = 2 * UserInteligencia

122         Case eClass.Assasin
124             Balance_AumentoMANA = UserInteligencia
                    
126         Case Else
128             Balance_AumentoMANA = 0
        
        
        
            
        End Select
        
        On Error GoTo Balance_AumentoMANA_Err

        Exit Function

Balance_AumentoMANA_Error:

130     LogError "Error " & Err.number & " (" & Err.description & ") in procedure Balance_AumentoMANA of Módulo mBalance in line " & Erl

        '<EhFooter>
        Exit Function

Balance_AumentoMANA_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.Balance_AumentoMANA " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Retorna la vida ideal que deberia tener el personaje para su nivel
Public Function getVidaIdeal(ByVal Elv As Byte, ByVal Class As Byte, ByVal Constitucion As Byte) As Single
        '<EhHeader>
        On Error GoTo getVidaIdeal_Err
        '</EhHeader>

        Dim promedio     As Single

        Dim vidaBase     As Integer

        Dim rangoAumento As tRango
    
100     vidaBase = 20
    
102     rangoAumento = getRangoAumentoVida(Class, Constitucion)
104     promedio = ((rangoAumento.minimo + rangoAumento.maximo) / 2)
    
106     getVidaIdeal = ((vidaBase + (Elv - 1) * promedio))

        '<EhFooter>
        Exit Function

getVidaIdeal_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.getVidaIdeal " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' El personaje aumenta su vida
Public Function obtenerAumentoHp(ByVal UserIndex As Integer) As Byte
        '<EhHeader>
        On Error GoTo obtenerAumentoHp_Err
        '</EhHeader>

        ' Calculo de vida
        Dim vidaPromedio  As Integer

        Dim promedio      As Single

        Dim vidaIdeal     As Single
        
        Dim minimoAumento As Integer

        Dim maximoAumento As Integer

        Dim aumentoHp     As Byte
    
        Dim rangoAumento  As tRango

        Dim vidaBase      As Integer
    
        Dim Random As Integer
    
100     vidaBase = 20

102     rangoAumento = getRangoAumentoVida(UserList(UserIndex).Clase, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
104     promedio = (rangoAumento.minimo + rangoAumento.maximo) / 2
106     vidaIdeal = vidaBase + (UserList(UserIndex).Stats.Elv - 2) * promedio
        
        Dim puntosAumento As Integer
        
120     puntosAumento = Int(RandomNumber(rangoAumento.minimo, rangoAumento.maximo))
                
122     If UserList(UserIndex).Stats.MaxHp < vidaIdeal + 1.5 Then
124        aumentoHp = maxi(puntosAumento, Int(0.5 + promedio))
        Else
            aumentoHp = puntosAumento
            
            If rangoAumento.minimo = puntosAumento Then
                If RandomNumber(1, 100) <= 50 Then
                    aumentoHp = puntosAumento + 1
                End If
            End If
172
        End If

    
174     obtenerAumentoHp = aumentoHp

        '<EhFooter>
        Exit Function

obtenerAumentoHp_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.obtenerAumentoHp " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub RecompensaPorNivel(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo RecompensaPorNivel_Err
        '</EhHeader>

        Dim Obj        As Obj

        Dim aumentoHp  As Single

        Dim AumentoMan As Integer

        Dim AumentoHit As Integer

        Dim AumentoSta As Integer
    
        Dim Texto      As String
    
        Dim Ups        As Single
    
100     With UserList(UserIndex)

102         If .Stats.Elv = STAT_MAXELV Then
104             .Stats.Exp = 0
106             .Stats.Elu = 0

            End If
        
108         Texto = "Nivel '" & .Stats.Elv & "' "

110         AumentoSta = Balance_AumentoSTA(.Clase)
112         aumentoHp = obtenerAumentoHp(UserIndex)
114         AumentoMan = Balance_AumentoMANA(.Clase, .Raza, .Stats.MaxMan)
116         AumentoHit = Balance_AumentoHIT(.Clase, .Stats.Elv)
        
118         .Stats.MaxHp = .Stats.MaxHp + aumentoHp
120         .Stats.MaxMan = .Stats.MaxMan + AumentoMan
122         .Stats.MinHit = .Stats.MinHit + AumentoHit
124         .Stats.MaxHit = .Stats.MaxHit + AumentoHit
126         .Stats.MaxSta = .Stats.MaxSta + AumentoSta
        
128         If .Stats.Elv < 36 Then
130             If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then .Stats.MinHit = STAT_MAXHIT_UNDER36
            Else

132             If .Stats.MinHit > STAT_MAXHIT_OVER36 Then .Stats.MinHit = STAT_MAXHIT_OVER36

            End If
        
134         If .Stats.Elv < 36 Then
136             If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then .Stats.MaxHit = STAT_MAXHIT_UNDER36
            Else

138             If .Stats.MaxHit > STAT_MAXHIT_OVER36 Then .Stats.MaxHit = STAT_MAXHIT_OVER36

            End If

140         Ups = .Stats.MaxHp - Mod_Balance.getVidaIdeal(.Stats.Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion))
            
142         Texto = Texto & "Vida +'" & aumentoHp & "' puntos de vida. Ups: " & IIf((Ups = 0), "Ninguno", Ups) & "."

144         If AumentoMan Then Texto = Texto & " Maná +" & AumentoMan
        
146         If AumentoHit Then
148             Texto = Texto & " Golpe +" & AumentoHit & ". "

            End If
        
150         Call WriteConsoleMsg(UserIndex, Texto, FontTypeNames.FONTTYPE_INFO)
        
152         Call Logs_User(.Name, eLog.eUser, eLvl, "paso a nivel " & .Stats.Elv & " gano HP: " & aumentoHp)
          
154         Call WriteUpdateUserStats(UserIndex)
    
        End With

        '<EhFooter>
        Exit Sub

RecompensaPorNivel_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.RecompensaPorNivel " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Reset_DesquiparAll(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Reset_DesquiparAll_Err
        '</EhHeader>

100     With UserList(UserIndex)

            'desequipar armadura
102         If .Invent.ArmourEqpObjIndex > 0 Then
104             Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

            End If
        
            ' Desequipamos la montura
106         If .Invent.MonturaObjIndex > 0 Then
108             Call Desequipar(UserIndex, .Invent.MonturaSlot)

            End If
        
            ' Desequipamos el pendiente
         If .Invent.PendientePartyObjIndex > 0 Then
             Call Desequipar(UserIndex, .Invent.PendientePartySlot)

            End If
            
            ' Desequipamos la reliquia
110         If .Invent.ReliquiaObjIndex > 0 Then
112             Call Desequipar(UserIndex, .Invent.ReliquiaSlot)

            End If
        
            ' Desequipamos el Objeto mágico (Laudes y Anillos mágicos)
114         If .Invent.MagicObjIndex > 0 Then
116             Call Desequipar(UserIndex, .Invent.MagicSlot)

            End If
        
            'desequipar arma
118         If .Invent.WeaponEqpObjIndex > 0 Then
120             Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

            End If
        
            'desequipar aura
122         If .Invent.AuraEqpObjIndex > 0 Then
124             Call Desequipar(UserIndex, .Invent.AuraEqpSlot)

            End If
        
            'desequipar casco
126         If .Invent.CascoEqpObjIndex > 0 Then
128             Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

            End If
        
            'desequipar herramienta
130         If .Invent.AnilloEqpSlot > 0 Then
132             Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

            End If
        
            'desequipar anillo magico/laud
134         If .Invent.MagicSlot > 0 Then
136             Call Desequipar(UserIndex, .Invent.MagicSlot)

            End If
        
            'desequipar municiones
138         If .Invent.MunicionEqpObjIndex > 0 Then
140             Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

            End If
        
            'desequipar escudo
142         If .Invent.EscudoEqpObjIndex > 0 Then
144             Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

            End If
    
        End With

        '<EhFooter>
        Exit Sub

Reset_DesquiparAll_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.Reset_DesquiparAll " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Obtiene el Promedio de Aumento de vida del Personaje
Public Function getPromedioAumentoVida(ByVal Class As Byte, ByVal Constitucion As Byte) As Single
        '<EhHeader>
        On Error GoTo getPromedioAumentoVida_Err
        '</EhHeader>

        Dim Rango As tRango
    
100     Rango = getRangoAumentoVida(Class, Constitucion)
    
102     getPromedioAumentoVida = (Rango.maximo + Rango.minimo) / 2

        '<EhFooter>
        Exit Function

getPromedioAumentoVida_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.getPromedioAumentoVida " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Retrona el minimo/maximo de puntos de vida que pude subir este usuario por nivel.
Public Function getRangoAumentoVida(ByVal Class As Byte, ByVal Constitucion As Byte) As tRango
        '<EhHeader>
        On Error GoTo getRangoAumentoVida_Err
        '</EhHeader>

100     getRangoAumentoVida.maximo = 0
102     getRangoAumentoVida.minimo = 0

104     Select Case Class

            Case eClass.Warrior

106             Select Case Constitucion

                    Case 21
108                     getRangoAumentoVida.minimo = 9
110                     getRangoAumentoVida.maximo = 12

112                 Case 20
114                     getRangoAumentoVida.minimo = 8
116                     getRangoAumentoVida.maximo = 12

118                 Case 19
120                     getRangoAumentoVida.minimo = 8
122                     getRangoAumentoVida.maximo = 11

124                 Case 18
126                     getRangoAumentoVida.minimo = 7
128                     getRangoAumentoVida.maximo = 11

130                 Case Else
132                     getRangoAumentoVida.minimo = 6 + AdicionalHPCazador
134                     getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

                End Select

136         Case eClass.Hunter
    
138             Select Case Constitucion

                    Case 21
140                     getRangoAumentoVida.minimo = 9
142                     getRangoAumentoVida.maximo = 11

144                 Case 20
146                     getRangoAumentoVida.minimo = 8
148                     getRangoAumentoVida.maximo = 11

150                 Case 19
152                     getRangoAumentoVida.minimo = 7
154                     getRangoAumentoVida.maximo = 11

156                 Case 18
158                     getRangoAumentoVida.minimo = 6
160                     getRangoAumentoVida.maximo = 11

162                 Case Else
164                     getRangoAumentoVida.minimo = 5
166                     getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

                End Select

168         Case eClass.Paladin

170             Select Case Constitucion

                    Case 21
172                     getRangoAumentoVida.minimo = 9
174                     getRangoAumentoVida.maximo = 11

176                 Case 20
178                     getRangoAumentoVida.minimo = 8
180                     getRangoAumentoVida.maximo = 11

182                 Case 19
184                     getRangoAumentoVida.minimo = 7
186                     getRangoAumentoVida.maximo = 11

188                 Case 18
190                     getRangoAumentoVida.minimo = 6
192                     getRangoAumentoVida.maximo = 11

194                 Case Else
196                     getRangoAumentoVida.minimo = 5
198                     getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

                End Select

200         Case eClass.Thief

202             Select Case Constitucion

                    Case 21
204                     getRangoAumentoVida.minimo = 6
206                     getRangoAumentoVida.maximo = 9

208                 Case 20
210                     getRangoAumentoVida.minimo = 5
212                     getRangoAumentoVida.maximo = 9

214                 Case 19
216                     getRangoAumentoVida.minimo = 4
218                     getRangoAumentoVida.maximo = 9

220                 Case 18
222                     getRangoAumentoVida.minimo = 4
224                     getRangoAumentoVida.maximo = 8

226                 Case 16, 17
228                     getRangoAumentoVida.minimo = 3
230                     getRangoAumentoVida.maximo = 7

232                 Case 16
234                     getRangoAumentoVida.minimo = 3
236                     getRangoAumentoVida.maximo = 6

238                 Case 14
240                     getRangoAumentoVida.minimo = 2
242                     getRangoAumentoVida.maximo = 6

244                 Case 13
246                     getRangoAumentoVida.minimo = 2
248                     getRangoAumentoVida.maximo = 5

250                 Case 12
252                     getRangoAumentoVida.minimo = 1
254                     getRangoAumentoVida.maximo = 5

256                 Case 11
258                     getRangoAumentoVida.minimo = 1
260                     getRangoAumentoVida.maximo = 4

262                 Case 10
264                     getRangoAumentoVida.minimo = 0
266                     getRangoAumentoVida.maximo = 4

268                 Case Else
270                     getRangoAumentoVida.minimo = 3
272                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPGuerrero

                End Select
    
274         Case eClass.Mage

276             Select Case Constitucion

                    Case 21
278                     getRangoAumentoVida.minimo = 6
280                     getRangoAumentoVida.maximo = 8

282                 Case 20
284                     getRangoAumentoVida.minimo = 5
286                     getRangoAumentoVida.maximo = 8

288                 Case 19
290                     getRangoAumentoVida.minimo = 4
292                     getRangoAumentoVida.maximo = 8

294                 Case 18
296                     getRangoAumentoVida.minimo = 3
298                     getRangoAumentoVida.maximo = 8

300                 Case Else
302                     getRangoAumentoVida.minimo = 3
304                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPGuerrero

                End Select
                
338         Case eClass.Cleric

340             Select Case Constitucion

                    Case 21
342                     getRangoAumentoVida.minimo = 7
344                     getRangoAumentoVida.maximo = 10

346                 Case 20
348                     getRangoAumentoVida.minimo = 6
350                     getRangoAumentoVida.maximo = 10

352                 Case 19
354                     getRangoAumentoVida.minimo = 6
356                     getRangoAumentoVida.maximo = 9

358                 Case 18
360                     getRangoAumentoVida.minimo = 5
362                     getRangoAumentoVida.maximo = 9

364                 Case Else
366                     getRangoAumentoVida.minimo = 4
368                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

                End Select

370         Case eClass.Druid

372             Select Case Constitucion

                    Case 21
374                     getRangoAumentoVida.minimo = 7
376                     getRangoAumentoVida.maximo = 10

378                 Case 20
380                     getRangoAumentoVida.minimo = 6
382                     getRangoAumentoVida.maximo = 10

384                 Case 19
386                     getRangoAumentoVida.minimo = 6
388                     getRangoAumentoVida.maximo = 9

390                 Case 18
392                     getRangoAumentoVida.minimo = 5
394                     getRangoAumentoVida.maximo = 9

396                 Case Else
398                     getRangoAumentoVida.minimo = 4
400                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

                End Select
        
402         Case eClass.Assasin

404             Select Case Constitucion

                    Case 21
406                     getRangoAumentoVida.minimo = 7
408                     getRangoAumentoVida.maximo = 10

410                 Case 20
412                     getRangoAumentoVida.minimo = 6
414                     getRangoAumentoVida.maximo = 10

416                 Case 19
418                     getRangoAumentoVida.minimo = 6
420                     getRangoAumentoVida.maximo = 9

422                 Case 18
424                     getRangoAumentoVida.minimo = 5
426                     getRangoAumentoVida.maximo = 9

428                 Case Else
430                     getRangoAumentoVida.minimo = 4
432                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

                End Select

434         Case eClass.Bard

436             Select Case Constitucion

                    Case 21
438                     getRangoAumentoVida.minimo = 7
440                     getRangoAumentoVida.maximo = 10

442                 Case 20
444                     getRangoAumentoVida.minimo = 6
446                     getRangoAumentoVida.maximo = 10

448                 Case 19
450                     getRangoAumentoVida.minimo = 6
452                     getRangoAumentoVida.maximo = 9

454                 Case 18
456                     getRangoAumentoVida.minimo = 5
458                     getRangoAumentoVida.maximo = 9

460                 Case Else
462                     getRangoAumentoVida.minimo = 4
464                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

                End Select


498         Case Else

500             Select Case Constitucion

                    Case 21
502                     getRangoAumentoVida.minimo = 6
504                     getRangoAumentoVida.maximo = 9

506                 Case 20
508                     getRangoAumentoVida.minimo = 5
510                     getRangoAumentoVida.maximo = 9

512                 Case 19
514                     getRangoAumentoVida.minimo = 4
516                     getRangoAumentoVida.maximo = 8

518                 Case Else
520                     getRangoAumentoVida.minimo = 5
522                     getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

                End Select

        End Select

        '<EhFooter>
        Exit Function

getRangoAumentoVida_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.Mod_Balance.getRangoAumentoVida " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

