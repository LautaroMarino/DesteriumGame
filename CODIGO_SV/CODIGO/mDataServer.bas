Attribute VB_Name = "mDataServer"
' Generamos los archivos que serán enviados al CLIENTE (Para evitar enviar datos por sockets)
Option Explicit

Public Type tCabeceraEncrypt

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type


Public CabeceraEncrypt    As tCabeceraEncrypt

Public Const PASSWD_CHARACTER As String = "AesirAO20TDSIMPERIUM"

Public Sub DataServer_Generate_ObjData()

        '<EhHeader>
        On Error GoTo DataServer_Generate_ObjData_Err

        '</EhHeader>
    
        Dim Manager  As clsIniManager

        Dim A        As Long, B As Long

        Dim FilePath As String

        Dim Anim     As Integer, AnimBajos As Integer

        Dim Temp     As String

100     FilePath = DatPath & "client\server_objs.ind"

        Set Manager = New clsIniManager
    
104     Call Manager.ChangeValue("INIT", "LASTOBJ", CStr(NumObjDatas))
    
106     For A = 1 To NumObjDatas

108         With ObjData(A)
110             Call Manager.ChangeValue(A, "NAME", mEncrypt_B.XOREncryption(.Name))
                Call Manager.ChangeValue(A, "GRHINDEX", CStr(.GrhIndex))
112             Call Manager.ChangeValue(A, "MINDEF", CStr(.MinDef))
114             Call Manager.ChangeValue(A, "MAXDEF", CStr(.MaxDef))
116             Call Manager.ChangeValue(A, "MINHIT", CStr(.MinHit))
118             Call Manager.ChangeValue(A, "MAXHIT", CStr(.MaxHit))
120             Call Manager.ChangeValue(A, "MINDEFRM", CStr(.DefensaMagicaMin))
122             Call Manager.ChangeValue(A, "MAXDEFRM", CStr(.DefensaMagicaMax))
                Call Manager.ChangeValue(A, "NOSECAE", CStr(.NoSeCae))
                Call Manager.ChangeValue(A, "OBJTYPE", CStr(.OBJType))
                Call Manager.ChangeValue(A, "VISUALSKIN", CStr(.VisualSkin))
                Call Manager.ChangeValue(A, "HOMBRE", CStr(.Hombre))
                Call Manager.ChangeValue(A, "MUJER", CStr(.Mujer))
                Call Manager.ChangeValue(A, "POINTS", CStr(.Points))
                
                ' Envia la animación correspondiente
                Select Case .OBJType
                    Case eOBJType.otarmadura, eOBJType.otTransformVIP
                        Anim = .Ropaje
                        AnimBajos = .RopajeEnano
                        
                    Case eOBJType.otcasco
                        Anim = .CascoAnim
                    
                    Case eOBJType.otescudo
                        Anim = .ShieldAnim
                    
                    Case eOBJType.otWeapon
                        Anim = .WeaponAnim
                        
                    Case eOBJType.otcofre
                        
                End Select
                
                Call Manager.ChangeValue(A, "ANIM", CStr(Anim))
                Call Manager.ChangeValue(A, "ANIMBAJOS", CStr(AnimBajos))
                Call Manager.ChangeValue(A, "PROYECTIL", CStr(.proyectil))
                Call Manager.ChangeValue(A, "DAMAGEMAG", CStr(.StaffDamageBonus))

                Call Manager.ChangeValue(A, "Skin", CStr(.Skin))
                Call Manager.ChangeValue(A, "GuildLvl", CStr(.GuildLvl))
                
                Call Manager.ChangeValue(A, "TIER", CStr(.Tier))
                Call Manager.ChangeValue(A, "VALUEDSP", CStr(.ValorEldhir))
                Call Manager.ChangeValue(A, "VALUEGLD", CStr(.valor))
                
                Call Manager.ChangeValue(A, "TIMEWARP", CStr(.TimeWarp))
                Call Manager.ChangeValue(A, "TIMEDURATION", CStr(.TimeDuration))
                Call Manager.ChangeValue(A, "REMOVEOBJ", CStr(.RemoveObj))
                Call Manager.ChangeValue(A, "PUEDEINSEGURA", CStr(.PuedeInsegura))
                
                Call Manager.ChangeValue(A, "LVLMIN", CStr(.LvlMin))
                Call Manager.ChangeValue(A, "LVLMAX", CStr(.LvlMax))
                
                If .SkillNum > 0 Then
                    Call Manager.ChangeValue(A, "SKILLS", CStr(.SkillNum))
                    
                    For B = 1 To .SkillNum
                        Call Manager.ChangeValue(A, "SK" & B, CStr(.Skill(B).Selected & "-" & .Skill(B).Amount))
                    Next B

                End If
               
                If .SkillsEspecialNum > 0 Then
                    Call Manager.ChangeValue(A, "SKILLSESPECIAL", CStr(.SkillsEspecialNum))
                    
                    For B = 1 To .SkillsEspecialNum
                        Call Manager.ChangeValue(A, "SKESP" & B, CStr(.SkillsEspecial(B).Selected & "-" & .SkillsEspecial(B).Amount))
                    Next B

                End If
                
                If .Upgrade.RequiredCant > 0 Then
                    Call Manager.ChangeValue(A, "REQUIREDCANT", CStr(.Upgrade.RequiredCant))
                
                    For B = 1 To .Upgrade.RequiredCant
                        Call Manager.ChangeValue(A, "R" & B, CStr(.Upgrade.Required(B).ObjIndex & "-" & .Upgrade.Required(B).Amount))
                    Next B

                End If
                
                For B = 1 To NUMCLASES

                    If ObjData(A).ClaseProhibida(B) <> 0 Then
                        Temp = Temp & ObjData(A).ClaseProhibida(B) & "-"

                    End If

                Next B
                
                If Temp <> vbNullString Then
                    Temp = Left$(Temp, Len(Temp) - 1)
                    Call Manager.ChangeValue(A, "CP", Temp)
                    Temp = vbNullString

                End If
                
                Call Manager.ChangeValue(A, "CHESTLAST", CStr(.Chest.NroDrop))
                

                If .Chest.NroDrop > 0 Then
                    Call Manager.ChangeValue(A, "PROBCLOSE", CStr(.Chest.ProbClose))
                    Call Manager.ChangeValue(A, "PROBBREAK", CStr(.Chest.ProbBreak))
                    Call Manager.ChangeValue(A, "RESPAWNTIME", CStr(.Chest.RespawnTime))
                    
                    For B = 1 To ObjData(A).Chest.NroDrop
                        Call Manager.ChangeValue(A, "CHEST" & B, .Chest.Drop(B))
                    Next B

                End If
                
            End With
        
            DoEvents
124     Next A
    
126     Call Manager.DumpFile(FilePath)
        Set Manager = Nothing
        '<EhFooter>
        Exit Sub

DataServer_Generate_ObjData_Err:
        LogError Err.description & vbCrLf & "in DataServer_Generate_ObjData " & "at line " & Erl
        Set Manager = Nothing

        '</EhFooter>
End Sub

' Procedimiento para agregar a la VisualSkin el objeto y así visualizarlo en la lista de objetos.
Public Sub DataServer_AddObjSkin(ByVal ObjIndex As Integer)

    With ObjData(ObjIndex)
        If .OBJType = eOBJType.otarmadura Or _
            .OBJType = eOBJType.otWeapon Or _
            .OBJType = eOBJType.otcasco Or _
            .OBJType = eOBJType.otescudo Then
        
        
            .VisualSkin = 1
        End If
    End With
    
    
End Sub

Public Sub DataServer_Generate_Npcs()

        '<EhHeader>
        On Error GoTo DataServer_Generate_Npcs_Err

        '</EhHeader>
    
        Dim Manager  As clsIniManager

        Dim N        As Integer

        Dim A        As Long, B As Long

        Dim NumNpcs  As Integer

        Dim FilePath As String
    
100     FilePath = DatPath & "client\server_npcs.ind"
102     Set Manager = New clsIniManager
   
104     NumNpcs = val(LeerNPCs.GetValue("INIT", "NUMNPCS"))
106     Call Manager.ChangeValue("INIT", "LASTNPC", CStr(NumNpcs))
    
108     For A = 1 To NumNpcs
110         Call Manager.ChangeValue(A, "NAME", mEncrypt_B.XOREncryption(LeerNPCs.GetValue("NPC" & A, "NAME")))
112         Call Manager.ChangeValue(A, "DESC", mEncrypt_B.XOREncryption(LeerNPCs.GetValue("NPC" & A, "DESC")))
            Call Manager.ChangeValue(A, "BODY", LeerNPCs.GetValue("NPC" & A, "BODY"))
            Call Manager.ChangeValue(A, "HEAD", LeerNPCs.GetValue("NPC" & A, "HEAD"))
            Call Manager.ChangeValue(A, "NPCTYPE", LeerNPCs.GetValue("NPC" & A, "NPCTYPE"))
            Call Manager.ChangeValue(A, "COMERCIA", LeerNPCs.GetValue("NPC" & A, "COMERCIA"))
            Call Manager.ChangeValue(A, "CRAFT", LeerNPCs.GetValue("NPC" & A, "QUEST"))
            
            Call Manager.ChangeValue(A, "EVASION", LeerNPCs.GetValue("NPC" & A, "PODEREVASION"))
            Call Manager.ChangeValue(A, "PODERATAQUE", LeerNPCs.GetValue("NPC" & A, "PODERATAQUE"))
            
            Call Manager.ChangeValue(A, "MINHIT", LeerNPCs.GetValue("NPC" & A, "MINHIT"))
            Call Manager.ChangeValue(A, "MAXHIT", LeerNPCs.GetValue("NPC" & A, "MAXHIT"))
            
            Call Manager.ChangeValue(A, "DEF", LeerNPCs.GetValue("NPC" & A, "DEF"))
            Call Manager.ChangeValue(A, "DEFM", LeerNPCs.GetValue("NPC" & A, "DEFM"))
            
            Call Manager.ChangeValue(A, "MAXHP", LeerNPCs.GetValue("NPC" & A, "MAXHP"))
            Call Manager.ChangeValue(A, "EXP", val(LeerNPCs.GetValue("NPC" & A, "GIVEEXP")) * MultExp)
            Call Manager.ChangeValue(A, "GLD", val(LeerNPCs.GetValue("NPC" & A, "GIVEGLD")) * MultGld)
            Call Manager.ChangeValue(A, "RESPAWNTIME", LeerNPCs.GetValue("NPC" & A, "RESPAWNTIME"))
            
            Dim NroItem As Byte

            Dim NroDrop As Byte

            Dim ln      As String
            
            NroItem = val(LeerNPCs.GetValue("NPC" & A, "NROITEMS"))
            
            Call Manager.ChangeValue(A, "NROITEMS", CStr(NroItem))
            
            For B = 1 To NroItem
                Call Manager.ChangeValue(A, "OBJ" & B, LeerNPCs.GetValue("NPC" & A, "Obj" & B))
            Next B
            
            NroDrop = val(LeerNPCs.GetValue("NPC" & A, "NRODROPS"))
            
            Call Manager.ChangeValue(A, "NRODROPS", CStr(NroDrop))
            
            For B = 1 To NroDrop
                Call Manager.ChangeValue(A, "DROP" & B, LeerNPCs.GetValue("NPC" & A, "DROP" & B))
            Next B
            
            DoEvents
114     Next A
    
116     Call Manager.DumpFile(FilePath)
    
118     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

DataServer_Generate_Npcs_Err:
        LogError Err.description & vbCrLf & "in DataServer_Generate_Npcs " & "at line " & Erl

        '</EhFooter>
End Sub

Public Sub DataServer_Generate_Quests()
        '<EhHeader>
        On Error GoTo DataServer_Generate_Quests_Err
        '</EhHeader>
    
        Dim Manager  As clsIniManager

        Dim N        As Integer

        Dim A        As Long, B As Long

        Dim NumNpcs  As Integer

        Dim FilePath As String
    
100     Set Manager = New clsIniManager
102     FilePath = DatPath & "client\server_quests.ind"
    
104     Call Manager.ChangeValue("INIT", "LASTQUEST", CStr(UBound(QuestList)))
    
106     For A = LBound(QuestList) To UBound(QuestList)

108         With QuestList(A)
110             Call Manager.ChangeValue(A, "NAME", mEncrypt_B.XOREncryption(.Nombre))
112             Call Manager.ChangeValue(A, "DESC", mEncrypt_B.XOREncryption(.Desc))
114             Call Manager.ChangeValue(A, "DESCFINAL", mEncrypt_B.XOREncryption(.DescFinish))
116             Call Manager.ChangeValue(A, "OBJ", CStr(.RequiredOBJs))
118             Call Manager.ChangeValue(A, "NPC", CStr(.RequiredNPCs))
120             Call Manager.ChangeValue(A, "REWARDOBJ", CStr(.RewardOBJs))
122             Call Manager.ChangeValue(A, "SALEOBJ", CStr(.RequiredSaleOBJs))
124             Call Manager.ChangeValue(A, "CHESTOBJ", CStr(.RequiredChestOBJs))
126             Call Manager.ChangeValue(A, "REWARDGLD", CStr(.RewardGLD))
128             Call Manager.ChangeValue(A, "REWARDEXP", CStr(.RewardEXP))
130             Call Manager.ChangeValue(A, "LASTQUEST", CStr(.LastQuest))
132             Call Manager.ChangeValue(A, "NEXTQUEST", CStr(.NextQuest))
134             Call Manager.ChangeValue(A, "REMOVE", CStr(.Remove))
            
136             For B = 1 To .RequiredOBJs
138                 Call Manager.ChangeValue(A, "OBJ" & B, .RequiredObj(B).ObjIndex & "-" & .RequiredObj(B).Amount)
140             Next B
        
142             For B = 1 To .RequiredSaleOBJs
144                 Call Manager.ChangeValue(A, "OBJSALE" & B, .RequiredSaleObj(B).ObjIndex & "-" & .RequiredSaleObj(B).Amount)
146             Next B

148             For B = 1 To .RequiredChestOBJs
150                 Call Manager.ChangeValue(A, "OBJCHEST" & B, .RequiredChestObj(B).ObjIndex & "-" & .RequiredChestObj(B).Amount)
152             Next B
            
154             For B = 1 To .RequiredNPCs
156                 Call Manager.ChangeValue(A, "NPC" & B, .RequiredNpc(B).NpcIndex & "-" & .RequiredNpc(B).Amount & "-" & .RequiredNpc(B).Hp)
158             Next B
            
160             For B = 1 To .RewardOBJs
162                 Call Manager.ChangeValue(A, "REWARDOBJ" & B, .RewardObj(B).ObjIndex & "-" & .RewardObj(B).Amount)
               
164             Next B

166             DoEvents
            
            End With

168     Next A
    
170     Call Manager.DumpFile(FilePath)
172     Set Manager = Nothing

        '<EhFooter>
        Exit Sub

DataServer_Generate_Quests_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mDataServer.DataServer_Generate_Quests " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DataServer_Generate_Shop()
        '<EhHeader>
        On Error GoTo DataServer_Generate_Shop_Err
        '</EhHeader>
    
        Dim Manager  As clsIniManager

        Dim A        As Long

        Dim FilePath As String

100     FilePath = DatPath & "client\server_shop.ind"

102     Set Manager = New clsIniManager
    
104     Call Manager.ChangeValue("INIT", "LAST", CStr(ShopLast))
    
106     For A = 1 To ShopLast

108         With Shop(A)
110             Call Manager.ChangeValue(A, "NAME", mEncrypt_B.XOREncryption(.Name))
112             Call Manager.ChangeValue(A, "DESC", mEncrypt_B.XOREncryption(.Desc))
                
114             Call Manager.ChangeValue(A, "GLD", CStr(.Gld))
116             Call Manager.ChangeValue(A, "DSP", CStr(.Dsp))
118             Call Manager.ChangeValue(A, "OBJINDEX", CStr(.ObjIndex) & "-" & CStr(.ObjAmount))
120             Call Manager.ChangeValue(A, "POINTS", CStr(.Points))
            
            End With
        
122         DoEvents
124     Next A
    
126     Call Manager.DumpFile(FilePath)
128     Set Manager = Nothing

        '<EhFooter>
        Exit Sub

DataServer_Generate_Shop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mDataServer.DataServer_Generate_Shop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub DataServer_Generate_Maps()
    
    Dim Manager  As clsIniManager

    Dim A        As Long, c As Long, B As Long

    Dim Text     As String
    
    Dim FilePath As String

    FilePath = DatPath & "client\server_maps.ind"

    Set Manager = New clsIniManager
    
    Call Manager.ChangeValue("INIT", "LAST", CStr(UBound(MiniMap)))
    
    For A = LBound(MiniMap) To UBound(MiniMap)

        With MiniMap(A)
            Call Manager.ChangeValue(A, "NAME", mEncrypt_B.XOREncryption(.Name))
            Call Manager.ChangeValue(A, "PK", CStr(.Pk))
                
            Call Manager.ChangeValue(A, "NPCSNUM", CStr(.NpcsNum))
                
            For B = 1 To .NpcsNum
                Call Manager.ChangeValue(A, "NPC_INDEX" & B, CStr(.Npcs(B).NpcIndex))
            Next B
            
            Call Manager.ChangeValue(A, "LVLMIN", CStr(.LvlMin))
            Call Manager.ChangeValue(A, "LVLMAX", CStr(.LvlMax))
            
            Call Manager.ChangeValue(A, "INVISINEFECTO", CStr(.InviSinEfecto))
            Call Manager.ChangeValue(A, "OCULTARSINEFECTO", CStr(.OcultarSinEfecto))
            Call Manager.ChangeValue(A, "RESUSINEFECTO", CStr(.ResuSinEfecto))
            Call Manager.ChangeValue(A, "INVOCARSINEFECTO", CStr(.InvocarSinEfecto))
            Call Manager.ChangeValue(A, "CAENITEM", CStr(.CaenItem))
                    
            Call Manager.ChangeValue(A, "SUB_MAPS", CStr(.Sub_Maps))
            Call Manager.ChangeValue(A, "CHESTLAST", CStr(.ChestLast))
            
            If .Sub_Maps > 0 Then

                For B = 1 To .Sub_Maps
                    Text = Text & .Maps(B) & "-"
                Next B
                
                Text = Left$(Text, Len(Text) - 1)
                
                Call Manager.ChangeValue(A, "MAPS", Text)
                Text = vbNullString

            End If
                
            If .ChestLast > 0 Then

                For B = 1 To .ChestLast
                    Text = Text & .Chest(B) & "-"
                Next B
                
                Text = Left$(Text, Len(Text) - 1)
                
                Call Manager.ChangeValue(A, "CHEST", Text)
                Text = vbNullString
                
            End If

        End With
        
        DoEvents
    Next A
    
    Call Manager.DumpFile(FilePath)
    Set Manager = Nothing

End Sub

Public Sub DataServer_Generate_Spells()
    
    Dim Manager  As clsIniManager

    Dim A        As Long, c As Long, B As Long

    Dim Text     As String
    
    Dim FilePath As String

    FilePath = DatPath & "client\server_spells.ind"

    Set Manager = New clsIniManager
    
    Call Manager.ChangeValue("INIT", "LAST", CStr(NumeroHechizos))
    
    For A = 1 To NumeroHechizos

        With Hechizos(A)
            Call Manager.ChangeValue(A, "NAME", mEncrypt_B.XOREncryption(.Nombre))
            Call Manager.ChangeValue(A, "AUTOLANZAR", CStr(.AutoLanzar))
            
        End With
        
        DoEvents
    Next A
    
    Call Manager.DumpFile(FilePath)
    Set Manager = Nothing

End Sub
