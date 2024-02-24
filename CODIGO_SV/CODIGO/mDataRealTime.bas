Attribute VB_Name = "mDataRealTime"
Option Explicit

Public Function Load_UserList_Offline(ByVal Name As String) As User

    On Error GoTo ErrHandler
    
    ' Cargamos el personaje
    Dim Userfile As clsIniManager
    Set Userfile = New clsIniManager
    
    Call Userfile.Initialize(CharPath & UCase$(Name) & ".chr")
    
    Dim TempUser As User
    Dim ln As String
    Dim LoopC As Long
    Dim A As Long
    Dim Temp As String
    
    With TempUser
        .Name = UCase$(Name)
        
        With .Stats
            .Points = CLng(Userfile.GetValue("STATS", "Points"))
            
            ' Skills
            For LoopC = 1 To NUMSKILLSESPECIAL
                .UserSkillsEspecial(LoopC) = val(Userfile.GetValue("SKILLSESPECIAL", "SKESP" & LoopC))
            Next LoopC
            
             For LoopC = 1 To NUMSKILLS
                .UserSkills(LoopC) = val(Userfile.GetValue("SKILLS", "SK" & LoopC))
                .EluSkills(LoopC) = val(Userfile.GetValue("SKILLS", "ELUSK" & LoopC))
                .ExpSkills(LoopC) = val(Userfile.GetValue("SKILLS", "EXPSK" & LoopC))
            Next LoopC
            
            ' # Hechizos
            For LoopC = 1 To MAXUSERHECHIZOS
                .UserHechizos(LoopC) = val(Userfile.GetValue("Hechizos", "H" & LoopC))
            Next LoopC
        
            .Eldhir = CLng(Userfile.GetValue("STATS", "ELDHIR"))
            .BonosHp = CLng(Userfile.GetValue("STATS", "BONOSHP"))
            .Gld = CLng(Userfile.GetValue("STATS", "GLD"))
            .MaxHp = CInt(Userfile.GetValue("STATS", "MaxHP"))
            .MinHp = CInt(Userfile.GetValue("STATS", "MinHP"))
            
            .SkillPts = CInt(Userfile.GetValue("STATS", "SkillPtsLibres"))
            .Exp = CDbl(Userfile.GetValue("STATS", "EXP"))
            .Elu = CLng(Userfile.GetValue("STATS", "ELU"))
            .Elv = CByte(Userfile.GetValue("STATS", "ELV"))
            
            
        
            ' # Bonificaciones del personaje
            .BonusLast = CInt(Userfile.GetValue("BONUS", "BONUSLAST"))
            
            If .BonusLast > 0 Then
                ReDim .Bonus(1 To .BonusLast) As UserBonus
                    
                For A = 1 To .BonusLast
                    Temp = Userfile.GetValue("BONUS", "BONUS" & A)
                    .Bonus(A).Tipo = val(ReadField(1, Temp, Asc("|")))
                    .Bonus(A).Value = val(ReadField(2, Temp, Asc("|")))
                    .Bonus(A).Amount = val(ReadField(3, Temp, Asc("|")))
                    .Bonus(A).DurationSeconds = val(ReadField(4, Temp, Asc("|")))
                    .Bonus(A).DurationDate = ReadField(5, Temp, Asc("|"))
                Next A
            End If
        End With
        
        ' # Faction
        With .Faction
            .FragsCiu = CLng(Userfile.GetValue("FACTION", "FragsCiu"))
            .FragsCri = CLng(Userfile.GetValue("FACTION", "FragsCri"))
            .FragsOther = CLng(Userfile.GetValue("FACTION", "FragsOther"))
            .Range = CByte(Userfile.GetValue("FACTION", "Range"))
            .Status = CByte(Userfile.GetValue("FACTION", "Status"))
            .StartDate = CStr(Userfile.GetValue("FACTION", "StartDate"))
            .StartElv = CByte(Userfile.GetValue("FACTION", "StartElv"))
            .StartFrags = CInt(Userfile.GetValue("FACTION", "StartFrags"))
            .ExFaction = CByte(Userfile.GetValue("FACTION", "ExFaction"))

        End With
            
        ' # Skins
         With .Skins
            .Last = CByte(Userfile.GetValue("SKINS", "LAST"))
            .ArmourIndex = CInt(Userfile.GetValue("SKINS", "ARMOUR"))
            .ShieldIndex = CInt(Userfile.GetValue("SKINS", "SHIELD"))
            .WeaponIndex = CInt(Userfile.GetValue("SKINS", "WEAPON"))
            .WeaponArcoIndex = CInt(Userfile.GetValue("SKINS", "WEAPONARCO"))
            .WeaponDagaIndex = CInt(Userfile.GetValue("SKINS", "WEAPONDAGA"))
            .HelmIndex = CInt(Userfile.GetValue("SKINS", "HELM"))
                
            ReDim .ObjIndex(1 To MAX_INVENTORY_SKINS) As Integer
                    
            For LoopC = 1 To MAX_INVENTORY_SKINS
                .ObjIndex(LoopC) = val(Userfile.GetValue("SKINS", CStr(LoopC)))
            Next LoopC

        End With
        
        ' # Bloqueo
        With .flags
            .Blocked = CByte(Userfile.GetValue("FLAGS", "BLOCKED"))
        End With
                
        ' # Apariencia y stats base
        .Genero = Userfile.GetValue("INIT", "Genero")
        .Clase = Userfile.GetValue("INIT", "Clase")
        .Raza = Userfile.GetValue("INIT", "Raza")
        .Hogar = Userfile.GetValue("INIT", "Hogar")
        .Char.Heading = CInt(Userfile.GetValue("INIT", "Heading"))
        
        With .OrigChar
            .Head = CInt(Userfile.GetValue("INIT", "Head"))
            .Body = CInt(Userfile.GetValue("INIT", "Body"))
            .WeaponAnim = CInt(Userfile.GetValue("INIT", "Arma"))
            .ShieldAnim = CInt(Userfile.GetValue("INIT", "Escudo"))
            .CascoAnim = CInt(Userfile.GetValue("INIT", "Casco"))
            
            .Heading = eHeading.SOUTH

        End With
        
        
        .Desc = Userfile.GetValue("INIT", "Desc")
        .Pos.Map = CInt(ReadField(1, Userfile.GetValue("INIT", "Position"), 45))
        .Pos.X = CInt(ReadField(2, Userfile.GetValue("INIT", "Position"), 45))
        .Pos.Y = CInt(ReadField(3, Userfile.GetValue("INIT", "Position"), 45))
        
        ' # Inventario-Banco
        .Invent.NroItems = CInt(Userfile.GetValue("Inventory", "CantidadItems"))
        .BancoInvent.NroItems = CInt(Userfile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = Userfile.GetValue("BancoInventory", "Obj" & LoopC)
            .BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
            .BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
        Next LoopC

        'Lista de objetos
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = Userfile.GetValue("Inventory", "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
        Next LoopC
         
        ' GuildIndex
        ln = Userfile.GetValue("Guild", "GUILDINDEX")

        If IsNumeric(ln) Then
            .GuildIndex = CInt(ln)
        Else
            .GuildIndex = 0
        End If
    End With
    
    
    Load_UserList_Offline = TempUser
    
    Exit Function
ErrHandler:
    Call LogError("Error en la funcion Load_UserList_Offline")
    
    
End Function
