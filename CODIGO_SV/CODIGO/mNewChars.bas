Attribute VB_Name = "mNewChars"
Option Explicit

Public Sub IUser_Editation_Skills(ByRef Temp As User)
        '<EhHeader>
        On Error GoTo IUser_Editation_Skills_Err
        '</EhHeader>
    
        ' Todas
100     Temp.Stats.UserSkills(eSkill.Resistencia) = LevelSkill(Temp.Stats.Elv).LevelValue
102     Temp.Stats.UserSkills(eSkill.Tacticas) = LevelSkill(Temp.Stats.Elv).LevelValue
104     Temp.Stats.UserSkills(eSkill.Armas) = LevelSkill(Temp.Stats.Elv).LevelValue
106     Temp.Stats.UserSkills(eSkill.Comerciar) = LevelSkill(Temp.Stats.Elv).LevelValue
108     Temp.Stats.UserSkills(eSkill.Navegacion) = 35
          Temp.Stats.UserSkills(eSkill.Ocultarse) = RandomNumber(1, 37)
           
        ' Clases con maná
110     If Temp.Stats.MaxMan > 0 Then
112         Temp.Stats.UserSkills(eSkill.Magia) = LevelSkill(Temp.Stats.Elv).LevelValue ' + RandomNumber(1, 7)
        End If
    
        ' Todas menos mago APUÑALAN
114     If Temp.Clase <> eClass.Mage Then
116           Temp.Stats.UserSkills(eSkill.Apuñalar) = LevelSkill(Temp.Stats.Elv).LevelValue
        End If
    
        ' Escudos
118     If Temp.Clase <> eClass.Mage And Temp.Clase <> eClass.Druid Then
120         Temp.Stats.UserSkills(eSkill.Defensa) = LevelSkill(Temp.Stats.Elv).LevelValue
        End If
    
        If Temp.Clase = eClass.Hunter Or Temp.Clase = eClass.Warrior Then
            Temp.Stats.UserSkills(eSkill.Proyectiles) = LevelSkill(Temp.Stats.Elv).LevelValue
        End If
        
        ' Armas de proyectiles
122     If Temp.Clase = eClass.Hunter Then
              Temp.Stats.UserSkills(eSkill.Ocultarse) = LevelSkill(Temp.Stats.Elv).LevelValue + RandomNumber(1, 14)
        End If
    
        ' Ocultarse & Robar
126     If Temp.Clase = eClass.Thief Then
128         Temp.Stats.UserSkills(eSkill.Ocultarse) = LevelSkill(Temp.Stats.Elv).LevelValue + RandomNumber(20, 40)
130         Temp.Stats.UserSkills(eSkill.Robar) = LevelSkill(Temp.Stats.Elv).LevelValue + RandomNumber(20, 40)
        End If
    
        ' Mineria, Tala, Pesca
134       Temp.Stats.UserSkills(eSkill.Pesca) = LevelSkill(Temp.Stats.Elv).LevelValue
136       Temp.Stats.UserSkills(eSkill.Mineria) = LevelSkill(Temp.Stats.Elv).LevelValue
138       Temp.Stats.UserSkills(eSkill.Talar) = LevelSkill(Temp.Stats.Elv).LevelValue
        
        
        
        
        ' Checking Final
        Dim A As Long
        
        For A = 1 To NUMSKILLS
        
            If Temp.Stats.UserSkills(A) > 100 Then
                Temp.Stats.UserSkills(A) = 100
            End If
            
        Next A
    
        '<EhFooter>
        Exit Sub

IUser_Editation_Skills_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.FrmPanelCreator.IUser_Editation_Skills " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub IUser_Editation_Reputacion_Frags(ByRef Temp As User, ByVal Frags As Integer)
        '<EhHeader>
        On Error GoTo IUser_Editation_Reputacion_Frags_Err
        '</EhHeader>
        Dim L     As Long

100     Frags = RandomNumber(val(FrmPanelCreator.txtFrags(0).Text), val(FrmPanelCreator.txtFrags(1).Text))
    
102     With Temp.Reputacion

104         If RandomNumber(1, 100) >= 50 Then
106             .AsesinoRep = (vlASESINO * 2)
108             If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
             
110              .BandidoRep = Frags + RandomNumber(1, 12) * 1000
             
112             .BurguesRep = 0
114             .NobleRep = 0
116             .PlebeRep = 0
118             Temp.Faction.FragsCiu = Frags
120             Temp.Faction.FragsCri = RandomNumber(1, Frags)
122             Temp.Faction.FragsOther = Temp.Faction.FragsCri + Temp.Faction.FragsCiu
            
            Else
124             .NobleRep = (vlNoble * Frags)

126             If .NobleRep > MAXREP Then .NobleRep = MAXREP
128             Temp.Faction.FragsCri = Frags
130             Temp.Faction.FragsOther = Temp.Faction.FragsCri
            End If
        
132         L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
134         L = L / 6
136         .promedio = L

        End With

        '<EhFooter>
        Exit Sub

IUser_Editation_Reputacion_Frags_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.FrmPanelCreator.IUser_Editation_Reputacion_Frags " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub IUser_Editation_Spells(ByRef Temp As User)

    Dim Slot As Byte
    
    With Temp
    
        Slot = 35
        .Stats.UserHechizos(Slot) = 10  ' Remover Parálisis
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 24  ' Inmovilizar
    
        If .Stats.MaxMan >= 1000 Then
            Slot = Slot - 1
            .Stats.UserHechizos(Slot) = 25  ' Apocalipsis

        End If
    
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 23  ' Descarga eléctrica
    
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 15  ' Tormenta de Fuego
    
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 18  ' Celeridad
    
        If .Clase <> eClass.Mage Then
            Slot = Slot - 1
            .Stats.UserHechizos(Slot) = 20  ' Fuerza

        End If
        
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 14  ' Invisibilidad
    
        If .Stats.MaxMan >= 1000 Then
            Slot = Slot - 1
            .Stats.UserHechizos(Slot) = 11  ' Resucitar

        End If
    
        Slot = Slot - 1
     
        If .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
            .Stats.UserHechizos(Slot) = 53  ' Invocar Mascotas
        Else
            .Stats.UserHechizos(Slot) = 52  ' Invocar Mascotas

        End If
    
        ' Hechizos básicos fuera de onda
        Slot = 1
        .Stats.UserHechizos(Slot) = 2  ' Curar Veneno
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 8  ' Misil Mágico
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 17  ' Invocar Zombies
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 3  ' Curar heridas leves
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 5  ' Curar heridas Graves
        
        If .Clase = eClass.Mage Then
            Slot = Slot + 1
            .Stats.UserHechizos(Slot) = 20  ' Fuerza

        End If
    
    End With

End Sub


