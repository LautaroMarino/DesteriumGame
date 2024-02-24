Attribute VB_Name = "mDataServer"
' Generamos los archivos que serán enviados al CLIENTE (Para evitar enviar datos por sockets)
Option Explicit


Public NpcsGlobal_Last As Integer
Public NpcsGlobal() As Integer

' Todo lo de empaquetado & Decript
Public Type tCabeceraEncrypt

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type


Public CabeceraEncrypt    As tCabeceraEncrypt

Public Const PASSWD_CHARACTER As String = "AesirAO20TDSIMPERIUM"


' Info skills atributes

Public Type eLevelSkill
    LevelValue As Integer
End Type

Public Type eInfoSkill
    Name As String
    MaxValue As Integer
    Color As Long
    bold As Boolean
End Type

Public LevelSkill(1 To 50)                            As eLevelSkill
Public InfoSkill(1 To NUMSKILLS)                        As eInfoSkill
Public InfoSkillEspecial(1 To NUMSKILLSESPECIAL)          As eInfoSkill

Public Type tRuletaItem

    ObjIndex As Integer      ' Index del objeto
    Amount As Integer       ' Cantidad de objeto que da
    Prob As Byte                ' 1,2,3,4,5
    ProbNum As Byte         ' 10,20,30,40,50,60,70,80,90a99

End Type

Public Type tRuletaConfig
    ItemLast As Integer
    Items() As tRuletaItem
    RuletaGld As Long
    RuletaDsp As Long
    

End Type

Public RuletaConfig As tRuletaConfig


Public Sub DataServer_LoadAll()
    
    Call DataServer_Load_ObjData
    Call DataServer_Generate_Npcs
    Call DataServer_Generate_Quests
    Call DataServer_Load_Shop
End Sub

Public Sub Determinate_Tier_Aura(ByVal ObjIndex As Integer)
    
End Sub
Public Sub DataServer_Load_ObjData()
    
    Dim Manager  As clsIniManager

    Dim A        As Long, b As Long

    Dim filePath As String

    Dim Temp     As String
    
    filePath = IniPath & "server\server_objs.ind"

    Set Manager = New clsIniManager
    
    Call Manager.Initialize(filePath)
    
    NumObjDatas = Val(Manager.GetValue("INIT", "LASTOBJ"))
    
    ReDim ObjData(1 To NumObjDatas) As tObjData
    ReDim CopyObjs(1 To NumObjDatas) As tObjData

    For A = 1 To NumObjDatas
        
        CopyObjs(A) = ObjData(A)
        
        With ObjData(A)
        
            .Name = mEncrypt_B.XORDecryption(Manager.GetValue(CStr(A), "NAME"))
            .GrhIndex = Val(Manager.GetValue(CStr(A), "GRHINDEX"))
            .MinDef = Val(Manager.GetValue(CStr(A), "MINDEF"))
            .MaxDef = Val(Manager.GetValue(CStr(A), "MAXDEF"))
            .MinHit = Val(Manager.GetValue(CStr(A), "MINHIT"))
            .MaxHit = Val(Manager.GetValue(CStr(A), "MAXHIT"))
            .MinDefRM = Val(Manager.GetValue(CStr(A), "MINDEFRM"))
            .MaxDefRM = Val(Manager.GetValue(CStr(A), "MAXDEFRM"))
            .ObjType = Val(Manager.GetValue(CStr(A), "OBJTYPE"))
            .Anim = Val(Manager.GetValue(CStr(A), "ANIM"))
            .AnimBajos = Val(Manager.GetValue(CStr(A), "ANIMBAJOS"))
            .Proyectil = Val(Manager.GetValue(CStr(A), "PROYECTIL"))
            .DamageMag = Val(Manager.GetValue(CStr(A), "DAMAGEMAG"))
            .ValueDSP = Val(Manager.GetValue(CStr(A), "VALUEDSP"))
            .Points = Val(Manager.GetValue(CStr(A), "POINTS"))
            .ValueGLD = Val(Manager.GetValue(CStr(A), "VALUEGLD"))
            .Skin = Val(Manager.GetValue(CStr(A), "SKIN"))
            .GuildLvl = Val(Manager.GetValue(CStr(A), "GUILDLVL"))
            .NoSeCae = Val(Manager.GetValue(CStr(A), "NOSECAE"))
            .RemoveObj = Val(Manager.GetValue(CStr(A), "REMOVEOBJ"))
            .TimeWarp = Val(Manager.GetValue(CStr(A), "TIMEWARP")) - 1
            .TimeDuration = Val(Manager.GetValue(CStr(A), "TIMEDURATION"))
            .PuedeInsegura = Val(Manager.GetValue(CStr(A), "PUEDEINSEGURA"))
            .VisualSkin = Val(Manager.GetValue(CStr(A), "VISUALSKIN"))
            .LvlMin = Val(Manager.GetValue(CStr(A), "LVLMIN"))
            .LvlMax = Val(Manager.GetValue(CStr(A), "LVLMAX"))
            .Tier = Val(Manager.GetValue(CStr(A), "TIER"))
            .Color = DataServer_ColorObj(.Tier)
            
            .Hombre = Val(Manager.GetValue(CStr(A), "HOMBRE"))
            .Mujer = Val(Manager.GetValue(CStr(A), "MUJER"))
            
            If .Skin > 0 Then
                SkinLast = SkinLast + 1
            End If

            .SkillNum = Val(Manager.GetValue(CStr(A), "SKILLS"))

            If .SkillNum > 0 Then
                ReDim .Skill(1 To .SkillNum) As ObjData_Skills
                
                For b = 1 To .SkillNum
                    Temp = Manager.GetValue(CStr(A), "SK" & b)
                    .Skill(b).Selected = Val(ReadField(1, Temp, 45))
                    .Skill(b).Amount = Val(ReadField(2, Temp, 45))
                Next b

            End If
            
            .SkillsEspecialNum = Val(Manager.GetValue(CStr(A), "SKILLSESPECIAL"))
            
            If .SkillsEspecialNum > 0 Then
                ReDim .SkillsEspecial(1 To .SkillsEspecialNum) As ObjData_Skills
                
                For b = 1 To .SkillsEspecialNum
                    Temp = Manager.GetValue(CStr(A), "SKESP" & b)
                    .SkillsEspecial(b).Selected = Val(ReadField(1, Temp, 45))
                    .SkillsEspecial(b).Amount = Val(ReadField(2, Temp, 45))
                Next b

            End If
            
            .Upgrade.RequiredCant = Val(Manager.GetValue(CStr(A), "REQUIREDCANT"))
            
            If .Upgrade.RequiredCant > 0 Then
                ReDim .Upgrade.Required(1 To .Upgrade.RequiredCant) As Obj
                
                For b = 1 To .Upgrade.RequiredCant
                     Temp = Manager.GetValue(CStr(A), "R" & b)
                    .Upgrade.Required(b).ObjIndex = Val(ReadField(1, Temp, 45))
                    .Upgrade.Required(b).Amount = Val(ReadField(2, Temp, 45))
                Next b

            End If

            Temp = Manager.GetValue(CStr(A), "CP")
            
            If Temp <> vbNullString Then
                .CP_Valid = True
                
                Dim CopyARR() As String
                CopyARR = Split(Temp, "-")
                
                ReDim .CP(LBound(CopyARR) To UBound(CopyARR)) As Byte
                For b = LBound(CopyARR) To UBound(CopyARR)
                    .CP(b) = Val(CopyARR(b))
                Next b
            End If
            
            .Chest.NroDrop = Val(Manager.GetValue(CStr(A), "CHESTLAST"))

            If .Chest.NroDrop > 0 Then
                .Chest.ProbClose = Val(Manager.GetValue(CStr(A), "PROBCLOSE"))
                .Chest.ProbBreak = Val(Manager.GetValue(CStr(A), "PROBBREAK"))
                .Chest.RespawnTime = Val(Manager.GetValue(CStr(A), "RESPAWNTIME"))
                ReDim .Chest.Drop(1 To .Chest.NroDrop) As Integer
                
                For b = 1 To .Chest.NroDrop
                    .Chest.Drop(b) = Val(Manager.GetValue(CStr(A), "CHEST" & b))
                Next b
            End If
            
            
            .ID = A
        End With
        
        DoEvents
    Next A
    
   ' Call Skins_Ordenate_ObjType(ObjData)
    
    Set Manager = Nothing

End Sub
' # Ordena los objetos para que aparezcan más juntos y luzca mejor.
Public Function Skins_Ordenate_ObjType(ByRef ObjData() As tObjData)

    Dim A    As Long, b As Long
    Dim Temp As tObjData
    
    CopyObjs = ObjData
    
    For A = 1 To NumObjDatas
        CopyObjs(A).ID = ObjData(A).ID
    Next A
    
    For A = 1 To NumObjDatas - 1
        For b = 1 To NumObjDatas - A

            With CopyObjs(b)
                    If .ObjType > CopyObjs(b + 1).ObjType Then
                        Temp = CopyObjs(b)
                        CopyObjs(b) = CopyObjs(b + 1)
                        CopyObjs(b + 1) = Temp
                        
                    
                    End If
            End With
        Next b
    Next A
                
End Function
Public Sub DataServer_Generate_Npcs()
        '<EhHeader>
        On Error GoTo DataServer_Generate_Npcs_Err
        '</EhHeader>
    
        Dim Manager  As clsIniManager
        Dim N        As Integer
        Dim A        As Long, b As Long, ln As String
        Dim filePath As String
    
100     filePath = IniPath & "server\server_npcs.ind"
102     Set Manager = New clsIniManager
        Call Manager.Initialize(filePath)
        
        NumNpcs = Val(Manager.GetValue("INIT", "LASTNPC"))
        ReDim NpcList(1 To NumNpcs) As tNpcs
        
108     For A = 1 To NumNpcs
            With NpcList(A)
                .Name = mEncrypt_B.XORDecryption(Manager.GetValue(A, "NAME"))
                .Desc = mEncrypt_B.XORDecryption(Manager.GetValue(A, "DESC"))
                .Body = Val(Manager.GetValue(A, "BODY"))
                .Head = Val(Manager.GetValue(A, "HEAD"))
                
                .NpcType = Val(Manager.GetValue(A, "NPCTYPE"))
                .Def = Val(Manager.GetValue(A, "DEF"))
                .DefM = Val(Manager.GetValue(A, "DEFM"))
                
                .MinHit = Val(Manager.GetValue(A, "MINHIT"))
                .MaxHit = Val(Manager.GetValue(A, "MAXHIT"))
                
                .MaxHp = Val(Manager.GetValue(A, "MAXHP"))
                .Comercia = Val(Manager.GetValue(A, "COMERCIA"))
                .Craft = Val(Manager.GetValue(A, "Craft"))
                .PoderEvasion = Val(Manager.GetValue(A, "PODEREVASION"))
                .PoderAtaque = Val(Manager.GetValue(A, "PODERATAQUE"))
                .GiveExp = Val(Manager.GetValue(A, "EXP"))
                .GiveGld = Val(Manager.GetValue(A, "GLD"))
                .RespawnTime = Val(Manager.GetValue(A, "RESPAWNTIME"))
                
                .NroItems = Val(Manager.GetValue(A, "NROITEMS"))
                .NroDrops = Val(Manager.GetValue(A, "NRODROPS"))
                
                For b = 1 To .NroItems
                    ln = Manager.GetValue(A, "OBJ" & b)
                    .Object(b).ObjIndex = Val(ReadField(1, ln, 45))
                    .Object(b).Amount = Val(ReadField(2, ln, 45))
                Next b
                
                For b = 1 To .NroDrops
                    ln = Manager.GetValue(A, "DROP" & b)
                    .Drop(b).ObjIndex = Val(ReadField(1, ln, 45))
                    .Drop(b).Amount = Val(ReadField(2, ln, 45))
                    .Drop(b).Probability = Val(ReadField(3, ln, 45))
                Next b
                
                DoEvents
            End With

114     Next A
    
118     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

DataServer_Generate_Npcs_Err:
        LogError err.Description & vbCrLf & _
           "in DataServer_Generate_Npcs " & _
           "at line " & Erl

        '</EhFooter>
End Sub

Public Sub DataServer_Generate_Quests()
    
    Dim Manager      As clsIniManager
    Dim N            As Integer
    Dim A            As Long, b As Long
    Dim filePath As String
    Dim Temp As String
    
    Set Manager = New clsIniManager
    filePath = IniPath & "server\server_quests.ind"
    Call Manager.Initialize(filePath)
        
    NumQuests = Val(Manager.GetValue("INIT", "LASTQUEST"))
    ReDim QuestList(1 To NumQuests) As tQuest
        
    For A = 1 To NumQuests
        With QuestList(A)
            .Name = mEncrypt_B.XORDecryption(Manager.GetValue(A, "NAME"))
            Temp = mEncrypt_B.XORDecryption(Manager.GetValue(A, "DESC"))
            .Desc = Split(Temp, "|")
            
            .DescFinish = mEncrypt_B.XORDecryption(Manager.GetValue(A, "DESCFINAL"))
            
            .RewardGld = Val(Manager.GetValue(A, "REWARDGLD"))
            .RewardExp = Val(Manager.GetValue(A, "REWARDEXP"))
                
            .Obj = Val(Manager.GetValue(A, "OBJ"))
            .Npc = Val(Manager.GetValue(A, "NPC"))
            .SaleObj = Val(Manager.GetValue(A, "SALEOBJ"))
            .ChestObj = Val(Manager.GetValue(A, "CHESTOBJ"))
                
            .RewardObj = Val(Manager.GetValue(A, "REWARDOBJ"))
            .LastQuest = Val(Manager.GetValue(A, "LASTQUEST"))
            .NextQuest = Val(Manager.GetValue(A, "NEXTQUEST"))
            .Remove = Val(Manager.GetValue(A, "REMOVE"))
                
            ReDim .Objs(0 To .Obj) As tObj_Quest
            ReDim .Npcs(0 To .Npc) As tNpc_Quest
            ReDim .RewardObjs(0 To .RewardObj) As tObj_Quest
            ReDim .SaleObjs(0 To .SaleObj) As tObj_Quest
            ReDim .ChestObjs(0 To .ChestObj) As tObj_Quest
            
            ReDim .NpcsUser(0 To .Npc) As tNpc_Quest
            ReDim .ObjsUser(0 To .Obj) As tObj_Quest
            ReDim .ObjsSaleUser(0 To .SaleObj) As tObj_Quest
            ReDim .ObjsChestUser(0 To .ChestObj) As tObj_Quest
    
            For b = 1 To .Obj
                Temp = Manager.GetValue(A, "OBJ" & b)
                .Objs(b).ObjIndex = ReadField(1, Temp, 45)
                .Objs(b).Amount = ReadField(2, Temp, 45)
            Next b
        
            For b = 1 To .SaleObj
                Temp = Manager.GetValue(A, "OBJSALE" & b)
                .SaleObjs(b).ObjIndex = ReadField(1, Temp, 45)
                .SaleObjs(b).Amount = ReadField(2, Temp, 45)
            Next b
            
            For b = 1 To .ChestObj
                Temp = Manager.GetValue(A, "OBJCHEST" & b)
                .ChestObjs(b).ObjIndex = ReadField(1, Temp, 45)
                .ChestObjs(b).Amount = ReadField(2, Temp, 45)
            Next b
            
            For b = 1 To .Npc
                 Temp = Manager.GetValue(A, "NPC" & b)
                .Npcs(b).NpcIndex = ReadField(1, Temp, 45)
                .Npcs(b).Amount = ReadField(2, Temp, 45)
                .Npcs(b).Hp = ReadField(3, Temp, 45)
            Next b
            
            For b = 1 To .RewardObj
                Temp = Manager.GetValue(A, "REWARDOBJ" & b)
                .RewardObjs(b).ObjIndex = ReadField(1, Temp, 45)
                .RewardObjs(b).Amount = ReadField(2, Temp, 45)
                '.RewardObjs(b).Durabilidad = ReadField(3, Temp, 45)
            Next b


            
        End With
            
        DoEvents
    Next A

    Set Manager = Nothing
End Sub


Public Sub DataServer_Load_Shop()
    
    Dim Manager      As clsIniManager
    Dim A            As Long
    Dim filePath As String
    Dim Temp As String
    
    filePath = IniPath & "server\server_shop.ind"

    Set Manager = New clsIniManager
    
    Call Manager.Initialize(filePath)
    
    ShopLast = Val(Manager.GetValue("INIT", "LAST"))
    
    ReDim Shop(1 To ShopLast) As tShop
    For A = 1 To ShopLast
        With Shop(A)
            .Name = mEncrypt_B.XORDecryption(Manager.GetValue(CStr(A), "NAME"))
             Temp = mEncrypt_B.XORDecryption(Manager.GetValue(CStr(A), "DESC"))
            .Desc = Split(Temp, "|")
            .Gld = Val(Manager.GetValue(CStr(A), "GLD"))
            .Dsp = Val(Manager.GetValue(CStr(A), "DSP"))
            
            Temp = Manager.GetValue(CStr(A), "OBJINDEX")
            .ObjIndex = Val(ReadField(1, Temp, 45))
            .ObjAmount = Val(ReadField(2, Temp, 45))
            .Points = Val(Manager.GetValue(CStr(A), "POINTS"))
            
        End With
        
        DoEvents
    Next A
    
    Set Manager = Nothing
End Sub

Public Sub DB_LoadSkills()
    Dim Manager As clsIniManager
    Dim A As Long
    Dim Temp As String
    Dim r As Byte, g As Byte, b As Byte
    Set Manager = New clsIniManager
        
        Manager.Initialize IniPath & "skills.ini"
        
        ' Skills por Nivel que puede ganar el personaje
        For A = 1 To 50
            LevelSkill(A).LevelValue = Val(Manager.GetValue("LEVELVALUE", "Lvl" & A))
        Next A
            
        ' Habilidades Cotidianas del Personaje
        For A = 1 To NUMSKILLS
            InfoSkill(A).Name = Manager.GetValue("SK" & A, "Name")
            InfoSkill(A).MaxValue = Val(Manager.GetValue("SK" & A, "MaxValue"))
            
            Temp = Manager.GetValue("SK" & A, "Color")
            r = Val(ReadField(1, Temp, 45))
            g = Val(ReadField(2, Temp, 45))
            b = Val(ReadField(3, Temp, 45))
            InfoSkill(A).Color = ARGB(r, g, b, 255)
            'InfoSkill(A).Bold = Manager.GetValue("SK" & A, "Bold")
        Next A
    
        ' Habilidades Extremas del Personaje
        For A = 1 To NUMSKILLSESPECIAL
            InfoSkillEspecial(A).Name = Manager.GetValue("SKESP" & A, "Name")
            InfoSkillEspecial(A).MaxValue = Val(Manager.GetValue("SKESP" & A, "MaxValue"))
            
            Temp = Manager.GetValue("SKESP" & A, "Color")
            r = Val(ReadField(1, Temp, 45))
            g = Val(ReadField(2, Temp, 45))
            b = Val(ReadField(3, Temp, 45))
            InfoSkillEspecial(A).Color = ARGB(r, g, b, 255)
            'InfoSkillEspecial(A).Bold = Manager.GetValue("SKESP" & A, "Bold")
        Next A
        
    Set Manager = Nothing
End Sub

' # Busca un slot repetido
Public Function ListNpc_Repeat(ByVal NpcIndex As Integer) As Boolean
    Dim A As Long
    
    If NpcsGlobal_Last = 0 Then Exit Function
    
    For A = 1 To NpcsGlobal_Last
        If NpcsGlobal(A) = NpcIndex Then
            ListNpc_Repeat = True
            Exit Function
        End If
    Next A
    
End Function
' # Guardamos el NPC en la lista global de npcs visibles para el usuario
Public Sub AddListNpcs(ByVal NpcIndex As Integer)
    
    If ListNpc_Repeat(NpcIndex) Then Exit Sub
    
    NpcsGlobal_Last = NpcsGlobal_Last + 1
    
    ReDim Preserve NpcsGlobal(1 To NpcsGlobal_Last) As Integer
    
    NpcsGlobal(NpcsGlobal_Last) = NpcIndex
End Sub

Public Sub DataServer_Load_Maps()
    
    Dim Manager  As clsIniManager

    Dim A        As Long, C As Long, b As Long

    Dim Text     As String
    
    Dim filePath As String

    filePath = IniPath & "server\server_maps.ind"

    Set Manager = New clsIniManager
    
    Manager.Initialize filePath
    
    MiniMap_Last = Val(Manager.GetValue("INIT", "LAST"))
    
    ReDim MiniMap(1 To MiniMap_Last) As tMinimap
    
    For A = LBound(MiniMap) To UBound(MiniMap)

        With MiniMap(A)
            .Name = mEncrypt_B.XORDecryption(Manager.GetValue(A, "NAME"))
            .Pk = Val(Manager.GetValue(A, "PK"))

            .NpcsNum = Val(Manager.GetValue(A, "NPCSNUM"))
            
            If .NpcsNum Then
                ReDim .Npcs(1 To .NpcsNum) As tMiniMap_Npc
                    
                For b = 1 To .NpcsNum
                    .Npcs(b).NpcIndex = Val(Manager.GetValue(A, "NPC_INDEX" & b))
                    
                    If NpcList(.Npcs(b).NpcIndex).MaxHp > 0 Then _
                    Call AddListNpcs(.Npcs(b).NpcIndex)
                Next b
                
                

            End If
            
            .LvlMin = Val(Manager.GetValue(A, "LVLMIN"))
            .LvlMax = Val(Manager.GetValue(A, "LVLMAX"))

            If .LvlMin = 0 Then .LvlMin = 1
            If .LvlMax = 0 Then .LvlMax = 47
                  
            .InviSinEfecto = Val(Manager.GetValue(A, "INVISINEFECTO"))
            .OcultarSinEfecto = Val(Manager.GetValue(A, "OCULTARSINEFECTO"))
            .ResuSinEfecto = Val(Manager.GetValue(A, "RESUSINEFECTO"))
            .InvocarSinEfecto = Val(Manager.GetValue(A, "INVOCARSINEFECTO"))
            .CaenItem = Val(Manager.GetValue(A, "CAENITEM"))
            .SubMaps = Val(Manager.GetValue(A, "SUB_MAPS"))
            .ChestLast = Val(Manager.GetValue(A, "CHESTLAST"))
                 
            If .SubMaps > 0 Then

                Dim ArraiMaps() As String
                    
                ArraiMaps = Split(Manager.GetValue(A, "MAPS"), "-")
                    
                ReDim .Maps(1 To .SubMaps) As Integer

                For b = 0 To .SubMaps - 1
                    .Maps(b + 1) = Val(ArraiMaps(b))
                Next b

            End If
                  
            If .ChestLast > 0 Then

                Dim ArraiChest() As String
                    
                ArraiChest = Split(Manager.GetValue(A, "CHEST"), "-")
                    
                ReDim .Chest(1 To .ChestLast) As Integer

                For b = 0 To .ChestLast - 1
                    .Chest(b + 1) = Val(ArraiChest(b))
                Next b

            End If

        End With
        
        DoEvents
    Next A
    
    Set Manager = Nothing

End Sub

Public Sub DataServer_Load_Spells()
    
    Dim Manager  As clsIniManager

    Dim A        As Long, C As Long, b As Long

    Dim Text     As String
    
    Dim filePath As String

    filePath = IniPath & "server\server_spells.ind"

    Set Manager = New clsIniManager
    
    Manager.Initialize filePath
    
    NumeroHechizos = Val(Manager.GetValue("INIT", "LAST"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    For A = 1 To NumeroHechizos

        With Hechizos(A)
            .Nombre = mEncrypt_B.XORDecryption(Manager.GetValue(A, "NAME"))
            .AutoLanzar = Val(Manager.GetValue(A, "AUTOLANZAR"))
        End With
        
        DoEvents
    Next A
    
    Set Manager = Nothing

End Sub
Public Sub Drops_Load()
        '<EhHeader>
        On Error GoTo Drops_Load_Err
        '</EhHeader>
        Dim Manager As clsIniManager
        Dim A As Long, b As Long
        Dim Temp As String
    
100     Set Manager = New clsIniManager

          Dim filePath As String
          filePath = Drops_FilePath
102     Manager.Initialize (filePath)
    
104     DropLast = Val(Manager.GetValue("INIT", "LAST"))
    
106     ReDim DropData(1 To DropLast) As tDrop
    
108     For A = 1 To DropLast
110         With DropData(A)
112             .Last = Val(Manager.GetValue(A, "LAST"))
            
114             ReDim .data(1 To .Last) As tDropData
            
116             For b = 1 To .Last
118                 Temp = Manager.GetValue(A, b)
120                 .data(b).ObjIndex = Val(ReadField(1, Temp, 45))
                      .data(b).Prob = Val(ReadField(2, Temp, 45))
122                 .data(b).Amount(0) = Val(ReadField(3, Temp, 45))
124                 .data(b).Amount(1) = Val(ReadField(4, Temp, 45))
126             Next b
        
            End With
    
    
128     Next A
130     Set Manager = Nothing
    
        '<EhFooter>
        Exit Sub

Drops_Load_Err:
        Set Manager = Nothing
        LogError err.Description & vbCrLf & _
               "in ServidorArgentum.mDrop.Drops_Load " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
Public Sub Ruleta_LoadItems()
        '<EhHeader>
        On Error GoTo Ruleta_LoadItems_Err
        '</EhHeader>

        Dim Manager As clsIniManager

        Dim A       As Long

        Dim Temp    As String
    
        Dim filePath As String
    
100     Set Manager = New clsIniManager
    
102     filePath = IniPath & "server\ruleta.dat"
            
            Manager.Initialize filePath
104     With RuletaConfig
106         .ItemLast = Val(Manager.GetValue("INIT", "LAST"))
108         .RuletaDsp = Val(Manager.GetValue("INIT", "RULETADSP"))
110         .RuletaGld = Val(Manager.GetValue("INIT", "RULETAGLD"))
        
112         If .ItemLast > 0 Then
114             ReDim .Items(1 To .ItemLast) As tRuletaItem
        
116             For A = 1 To .ItemLast

118                 With .Items(A)
120                     Temp = Manager.GetValue("LIST", "OBJ" & A)
                
122                     .ObjIndex = Val(ReadField(1, Temp, 45))
124                     .Amount = Val(ReadField(2, Temp, 45))
126                     .Prob = Val(ReadField(3, Temp, 45))
128                     .ProbNum = Val(ReadField(4, Temp, 45))
                
                    End With

130             Next A
    
            End If
    
        End With
    
132     Set Manager = Nothing

        '<EhFooter>
        Exit Sub

Ruleta_LoadItems_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.mDataServer.Ruleta_LoadItems " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


