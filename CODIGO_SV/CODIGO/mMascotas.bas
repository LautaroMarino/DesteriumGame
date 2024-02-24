Attribute VB_Name = "mMascotas"
Option Explicit

Private Const ARCHIVE As String = "MASCOTAS.DAT"

Private Type tMascota

    Name As String
    MinHit As Integer
    MaxHit As Integer
    MinHitMag As Integer
    MaxHitMag As Integer
    Spells(1 To 35) As Integer
    
    SoloMagia As Boolean
    SoloGolpe As Boolean

End Type

Public Mascotas() As tMascota

Public Function Mascota_Index(ByVal UserIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo Mascota_Index_Err
        '</EhHeader>

100     With UserList(UserIndex)
            'Druidas
102         If .Clase = eClass.Druid Then

104             Select Case .Raza

                    Case eRaza.Humano, eRaza.Gnomo, eRaza.Enano
106                     Mascota_Index = 78

                        Exit Function

108                 Case eRaza.Elfo, eRaza.Drow
110                     Mascota_Index = 96

                        Exit Function

                End Select

            End If
        
            ' Clerigos
112         If .Clase = eClass.Cleric Then
114             Mascota_Index = 92

                Exit Function

            End If
        
            ' Bardos
116         If .Clase = eClass.Bard Then
118             Mascota_Index = 94

                Exit Function

            End If
        
            ' Magos
120         If .Clase = eClass.Mage Then
122             Mascota_Index = 93

                Exit Function

            End If
        
            ' Paladines-Asesinos
124         If .Clase = eClass.Assasin Or .Clase = eClass.Paladin Then
126             Mascota_Index = 115
            End If
    
        End With
    
        '<EhFooter>
        Exit Function

Mascota_Index_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMascotas.Mascota_Index " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Mascotas_AddNew(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo Mascotas_AddNew_Err
        '</EhHeader>
    
        Dim Slot As Byte

        Dim Obj  As Obj

100     With UserList(UserIndex)
        
            'Slot = Mascota_FreeSlot(UserIndex)
        
102         WriteConsoleMsg UserIndex, "Sabemos lo importante que es este sistema para vos. Te prometemos un nuevo sistema de domar mascotas, donde podrás entrenarlas. Estamos trabajando en ello. Mientras tanto tendrás los hechizos para invocar mascotas momentaneas.", FontTypeNames.FONTTYPE_INFO

            Exit Sub
        
104         If Slot = 0 Then
                'WriteConsoleMsg UserIndex, "No tienes lugar para mas mascotas.", FontTypeNames.FONTTYPE_INFO
                'Exit Sub
            End If
        
106         If RandomNumber(1, 100) <= 77 Then
108             WriteConsoleMsg UserIndex, "No has logrado domar a la criatura.", FontTypeNames.FONTTYPE_INFO

                Exit Sub

            End If
        
110         Obj.ObjIndex = Npclist(NpcIndex).MonturaIndex
112         Obj.Amount = 1
        
114         If Not MeterItemEnInventario(UserIndex, Obj) Then
116             WriteConsoleMsg UserIndex, "No tienes lugar en tu inventario. SI o SI debes tenerla en él.", FontTypeNames.FONTTYPE_INFO

                Exit Sub

            End If
        
        End With

        '<EhFooter>
        Exit Sub

Mascotas_AddNew_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMascotas.Mascotas_AddNew " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DoEquita(ByVal UserIndex As Integer, _
                    ByRef Montura As ObjData, _
                    ByVal Slot As Integer)

        '<EhHeader>
        On Error GoTo DoEquita_Err

        '</EhHeader>

100     With UserList(UserIndex)
        
102         If .flags.Montando = 0 Then
104             .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
106             .Invent.MonturaSlot = Slot
108             .Char.Head = 0

110             If .flags.Muerto = 0 Then
112                 .Char.Body = Montura.Ropaje
                Else
114                 .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
116                 .Char.Head = iCabezaMuerto(Escriminal(UserIndex))

                End If

118             .Char.Head = UserList(UserIndex).OrigChar.Head
120             .Char.ShieldAnim = NingunEscudo
122             .Char.WeaponAnim = NingunArma
124             .Char.CascoAnim = .Char.CascoAnim
126             .flags.Montando = 1
128             .Invent.Object(Slot).Equipped = 1
            Else

130             If .Invent.MonturaObjIndex <> .Invent.Object(Slot).ObjIndex Then
132                 Call WriteConsoleMsg(UserIndex, "Esta no es la montura a la que estabas subido.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
            
134             .Invent.Object(Slot).Equipped = 0
136             .flags.Montando = 0
            
138             If .flags.Muerto = 0 Then
140                 .Char.Head = UserList(UserIndex).OrigChar.Head

142                 If .Invent.ArmourEqpObjIndex > 0 Then
144                     .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
                    Else
146                     Call DarCuerpoDesnudo(UserIndex)

                    End If

148                 If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
150                 If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
152                 If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
                Else
154                 .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
156                 .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
                    
158                 .Char.ShieldAnim = NingunEscudo
160                 .Char.WeaponAnim = NingunArma
162                 .Char.CascoAnim = NingunCasco
                      
                      Dim A As Long
                      
                      For A = 1 To MAX_AURAS
164                     .Char.AuraIndex(A) = NingunAura
                     Next A
                End If

            End If
      
166         Call WriteChangeInventorySlot(UserIndex, Slot)
168         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
170         Call WriteMontateToggle(UserIndex)

        End With

        '<EhFooter>
        Exit Sub

DoEquita_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMascotas.DoEquita " & "at line " & Erl

        

        '</EhFooter>
End Sub
