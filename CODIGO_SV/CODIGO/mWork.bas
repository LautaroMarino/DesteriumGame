Attribute VB_Name = "mWork"
' Módulo de trabajo DesteriumAO Exodo III

Option Explicit

Enum eItemsConstruibles_Subtipo
    eArmadura = 1
    eCasco = 2
    eEscudo = 3
    eArmas = 4
    eMuniciones = 5
    eEmbarcaciones = 6
    eObjetoMagico = 7
    eInstrumento = 8
End Enum

Public Type eItemsConstruibles
    ObjIndex As Integer
    SubTipo As eItemsConstruibles_Subtipo
End Type

Public ObjBlacksmith() As eItemsConstruibles
Public ObjBlacksmith_Amount As Integer

Public ObjCarpinter() As Integer
Public ObjCarpinter_Amount As Integer

Public Sub Crafting_Reset()
    ObjBlacksmith_Amount = 0
    ReDim ObjBlacksmith(0) As eItemsConstruibles
End Sub
' # Agregamos el objeto a la lista de herrería
Public Sub Crafting_BlackSmith_Add(ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo Crafting_BlackSmith_Add_Err
        '</EhHeader>

100     ObjBlacksmith_Amount = ObjBlacksmith_Amount + 1
102     ReDim Preserve ObjBlacksmith(0 To ObjBlacksmith_Amount) As eItemsConstruibles
    
104     ObjBlacksmith(ObjBlacksmith_Amount).ObjIndex = ObjIndex
106     ObjBlacksmith(ObjBlacksmith_Amount).SubTipo = Set_Subtype_Object(ObjIndex)
    
    
        '<EhFooter>
        Exit Sub

Crafting_BlackSmith_Add_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mWork.Crafting_BlackSmith_Add " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' # Agregamos el SubTipo del objeto, para luego separarlos de forma mas rapida
Public Function Set_Subtype_Object(ByVal ObjIndex As Integer) As eItemsConstruibles_Subtipo
        '<EhHeader>
        On Error GoTo Set_Subtype_Object_Err
        '</EhHeader>
    
100     With ObjData(ObjIndex)
        
102         Select Case .OBJType
        
                Case eOBJType.otarmadura
104                 Set_Subtype_Object = eArmadura
                
106             Case eOBJType.otAnillo, eOBJType.otMagic, eOBJType.oteffect
108                 Set_Subtype_Object = eObjetoMagico
                
110             Case eOBJType.otescudo
112                 Set_Subtype_Object = eEscudo
                
114             Case eOBJType.otcasco
116                 Set_Subtype_Object = eCasco
                
118             Case eOBJType.otWeapon
120                 Set_Subtype_Object = eArmas
                
122             Case eOBJType.otFlechas
124                 Set_Subtype_Object = eMuniciones
                
126             Case eOBJType.otBarcos
128                 Set_Subtype_Object = eEmbarcaciones
                
130             Case eOBJType.otInstrumentos
132                 Set_Subtype_Object = eInstrumento
                
134             Case Else
136                 Set_Subtype_Object = eObjetoMagico
            End Select
        
        End With

        '<EhFooter>
        Exit Function

Set_Subtype_Object_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mWork.Set_Subtype_Object " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' # Comprueba de tener los recursos necesarios que necesita el objeto para ser creado/mejorado
Public Function Crafting_Checking_Object(ByVal UserIndex As Integer, _
                                    ByVal QuestIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo Crafting_Checking_Object_Err
        '</EhHeader>
        Dim A As Long
        Dim Temp As String
    
100     Crafting_Checking_Object = True
    
102     With QuestList(QuestIndex)
104         For A = 1 To .RequiredOBJs
106             If Not TieneObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex) Then
               
108                 Crafting_Checking_Object = False
                    Exit Function
                End If
110         Next A
        
        End With
    
    
        '<EhFooter>
        Exit Function

Crafting_Checking_Object_Err:
        LogError Err.description & vbCrLf & _
               "in Crafting_Checking_Object " & _
               "at line " & Erl

        '</EhFooter>
End Function

' # Quita los recursos necesarios para la construcción/mejora del objeto.
Public Sub Crafting_Remove_Object(ByVal UserIndex As Integer, _
                                  ByVal QuestIndex As Integer)
        '<EhHeader>
        On Error GoTo Crafting_Remove_Object_Err
        '</EhHeader>
    
        Dim A As Long
    
100     With QuestList(QuestIndex)
102         For A = 1 To .RequiredOBJs
104             Call QuitarObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex)
106         Next A
        End With
    
        '<EhFooter>
        Exit Sub

Crafting_Remove_Object_Err:
        LogError Err.description & vbCrLf & _
               "in Crafting_Remove_Object " & _
               "at line " & Erl

        '</EhFooter>
End Sub



' # Fundicion del Objeto
Private Sub Crafting_Fundition(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo Crafting_Fundition_Err
        '</EhHeader>
    
        Dim A As Long
        Dim Temp As Long
        Dim Obj As Obj
    
100     If ConfigServer.ModoCrafting = 0 Then
102         Call WriteConsoleMsg(UserIndex, "El servidor no admite la fundición de objetos.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        End If
    
104     With UserList(UserIndex)
        
        
106         For A = 1 To ObjData(ObjIndex).Upgrade.RequiredCant
108             Temp = ObjData(ObjIndex).Upgrade.Required(A).Amount * 0.3
            
110             If Temp > 0 Then
112                 Obj.Amount = Temp
114                 Obj.ObjIndex = ObjData(ObjIndex).Upgrade.Required(A).ObjIndex
                
116                 If Not MeterItemEnInventario(UserIndex, Obj) Then
118                     Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj)
                    End If
                End If
120         Next A
    
122         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sConstruction, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
124         UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta

126         If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        End With
        '<EhFooter>
        Exit Sub

Crafting_Fundition_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mWork.Crafting_Fundition " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

