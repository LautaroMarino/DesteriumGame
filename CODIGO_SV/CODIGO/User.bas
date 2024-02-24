Attribute VB_Name = "UserNew"
Option Explicit


' Creo un nuevo personaje
Public Sub User_Create(ByVal UserIndex As Integer, ByRef Class As eClass)
        
    Dim Char As Char
    
    With UserList(UserIndex)
        .Clase = Class
        .Raza = Balance.RazeClass(.Clase)
        .Genero = Balance.GeneroClass(.Clase)
        
        Call User_Set_Stats(.Clase, .Stats)
        
        Char = User_Set_Apparience(.Raza, .Genero, .Char)
        
        Call User_Set_Inventory(UserList(UserIndex))
        
        .OrigChar = .Char
    End With
        
End Sub

' Setea el Inventario PRE cargado
Public Sub User_Set_Inventory(ByRef UserI As User)
    
    Dim Slot As Byte
    Dim A As Long
    Dim ObjIndex As Integer
    
    Dim BalanceTemp As tBalance_ClassObj
    BalanceTemp = Balance.ListObjs(UserI.Clase)
    
    With UserI
        For A = LBound(BalanceTemp.Obj()) To UBound(BalanceTemp.Obj())
            ObjIndex = BalanceTemp.Obj(A)
            
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = ObjIndex
            .Invent.Object(Slot).Amount = 1
            
            Select Case ObjData(ObjIndex).OBJType
            
                Case eOBJType.otarmadura
                    .Invent.ArmourEqpObjIndex = ObjIndex
                    .Invent.ArmourEqpSlot = Slot
                    .Invent.Object(Slot).Equipped = 1
                    .flags.Desnudo = 0
                    .Char.Body = GetArmourAnim(UserI.Raza, .Invent.ArmourEqpObjIndex)
                Case eOBJType.otWeapon
                    .Invent.WeaponEqpObjIndex = ObjIndex
                    .Invent.WeaponEqpSlot = Slot
                    .Invent.Object(Slot).Equipped = 1
                    .Char.WeaponAnim = GetWeaponAnim(UserI.Raza, ObjIndex)
                Case eOBJType.otescudo
                    .Invent.EscudoEqpObjIndex = ObjIndex
                    .Invent.EscudoEqpSlot = Slot
                    .Invent.Object(Slot).Equipped = 1
                    .Char.ShieldAnim = ObjData(ObjIndex).ShieldAnim
                Case eOBJType.otcasco
                    .Invent.CascoEqpObjIndex = ObjIndex
                    .Invent.CascoEqpSlot = Slot
                    .Invent.Object(Slot).Equipped = 1
                    .Char.CascoAnim = ObjData(ObjIndex).CascoAnim
            End Select
        Next A
        
        ' Total Items
         .Invent.NroItems = Slot
    End With

End Sub

' Apariencia del Personaje
Public Function User_Set_Apparience(ByRef Raze As eRaza, ByVal Genero As Byte, ByRef CharA As Char) As Char

    With CharA
        .Heading = eHeading.SOUTH
        .Head = IHead_Generate(Genero, Raze)
        .Body = IBody_Generate(Genero, Raze)
        
        .ShieldAnim = NingunEscudo
        .CascoAnim = NingunCasco
        .WeaponAnim = NingunArma
    
    End With

End Function

' Inicia los atributos
Public Sub User_Set_Stats(ByRef Class As eClass, ByRef Stats As UserStats)
    
    With Stats
        .MaxHp = Balance.Health_Initial(Class)
        .MaxMan = Balance.Mana_Initial(Class)
        .Armour = Balance.Armour_Initial(Class)
        .ArmourMag = Balance.ArmourMag_Initial(Class)
        .Damage = Balance.Damage_Initial(Class)
        .DamageMag = Balance.DamageMag_Initial(Class)
        .RegHP = Balance.RegHP_Initial(Class)
        .RegMANA = Balance.RegMANA_Initial(Class)
        .Cooldown = Balance.Cooldown_Initial(Class)
        .Movement = Balance.Movement_Initial(Class)
        
        .MinHp = .MaxHp
        .MinMan = .MaxMan
    End With
    
End Sub

' Mejora los atributos por Nivel
Public Sub User_Upgrade_Stats_Level(ByRef Class As eClass, ByRef Stats As UserStats)

    With Stats
        .MaxHp = .MaxHp + Balance.Health_Level(Class)
        .MaxMan = .MaxMan + Balance.Mana_Level(Class)
        .Armour = .Armour + Balance.Armour_Level(Class)
        .ArmourMag = .ArmourMag + Balance.ArmourMag_Level(Class)
        .Damage = .Damage + Balance.Damage_Level(Class)
        .DamageMag = .DamageMag + Balance.DamageMag_Level(Class)
        .RegHP = .RegHP + Balance.RegHP_Level(Class)
        .RegMANA = .RegMANA + Balance.RegMANA_Level(Class)
        .Cooldown = .Cooldown + Balance.Cooldown_Level(Class)
        
        .MinHp = .MaxHp
        .MinMan = .MaxMan
    End With
End Sub
