Attribute VB_Name = "InvNpc"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, _
                                Obj As Obj, _
                                Optional NotPirata As Boolean = True) As WorldPos
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo TirarItemAlPiso_Err
        '</EhHeader>



        Dim NuevaPos As WorldPos

100     NuevaPos.X = 0
102     NuevaPos.Y = 0
    
104     Tilelibre Pos, NuevaPos, Obj, NotPirata, True

106     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
108         Call MakeObj(Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
            'Pos = NuevaPos
        End If
    
110     TirarItemAlPiso = NuevaPos


        '<EhFooter>
        Exit Function

TirarItemAlPiso_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.TirarItemAlPiso " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' # Genera los DROPS para el DISCORD
Public Function NPC_LISTAR_ITEMS(ByVal NpcIndex As Integer) As String
    Dim Temp As String
    Dim A As Long
    Dim Porc As Double

    With Npclist(NpcIndex)
        
        Debug.Print .numero
        
        If .Invent.NroItems > 0 Then
            For A = 1 To .Invent.NroItems
                Temp = Temp & "**" & ObjData(.Invent.Object(A).ObjIndex).Name & "** (x" & .Invent.Object(A).Amount & ") **[100%]**"
     
                If A < .Invent.NroItems Then
                    Temp = Temp & " | "
                End If
            Next A
        
            
            Temp = Temp & vbCrLf
        End If
        
        For A = 1 To .NroDrops
            Porc = (.Drop(A).ProbNum / 10 ^ .Drop(A).Probability) * 10

            Temp = Temp & "**" & ObjData(.Drop(A).ObjIndex).Name & "** (x" & .Drop(A).Amount & ") **[" & Porc & "%]**"

            If A < .NroDrops Then
                Temp = Temp & " | "
            End If
        Next A
    End With

    NPC_LISTAR_ITEMS = Temp
End Function




Public Sub NPC_TIRAR_ITEMS(ByVal UserIndex As Integer, _
                           ByRef Npc As Npc, _
                           ByVal IsPretoriano As Boolean)
        '<EhHeader>
        On Error GoTo NPC_TIRAR_ITEMS_Err
        '</EhHeader>

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 28/11/2009
        'Give away npc's items.
        '28/11/2009: ZaMa - Implementado drops complejos
        '02/04/2010: ZaMa - Los pretos vuelven a tirar oro.
        '10/04/2011: ZaMa - Logueo los objetos logueables dropeados.
        '***************************************************


100     With Npc
        
            Dim A As Long, B As Long
            Dim Probability As Long
        
            Dim i        As Byte

            Dim MiObj    As Obj

            Dim NroDrop  As Integer

            Dim Random   As Integer

            Dim ObjIndex As Integer
            ' Si esta en party realiza la entrega por otro lado..
            
            If UserList(UserIndex).GroupIndex = 0 Then
                ' Dropea oro?
102             If .GiveGLD > 0 Then Call TirarOroNpc(UserIndex, .GiveGLD, .Pos)
            End If
            
            ' ¿Tiene objetos del inventario para tirar o Drops ?
104         For A = 1 To MAX_INVENTORY_SLOTS

106             If .Invent.Object(A).ObjIndex > 0 Then
108                 MiObj.Amount = .Invent.Object(A).Amount
110                 MiObj.ObjIndex = .Invent.Object(A).ObjIndex
112                 Call TirarItemAlPiso(.Pos, MiObj)
                End If
              Next A
              
              For A = 1 To .NroDrops
114             If .Drop(A).ObjIndex > 0 Then
116                 For B = 1 To .Drop(A).Probability

118                     If RandomNumber(1, 100) <= .Drop(A).ProbNum Then
120                         Probability = Probability + 1
                        End If
122                 Next B
                
124                 If Probability = .Drop(A).Probability Then
126                     MiObj.Amount = .Drop(A).Amount
128                     MiObj.ObjIndex = .Drop(A).ObjIndex
130                     Call TirarItemAlPiso(.Pos, MiObj)
                          Exit Sub
                    End If
                
132                 Probability = 0
                End If
            

134         Next A

        End With
    


        '<EhFooter>
        Exit Sub

NPC_TIRAR_ITEMS_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.NPC_TIRAR_ITEMS " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo QuedanItems_Err
        '</EhHeader>

        Dim i As Integer

100     If Npclist(NpcIndex).Invent.NroItems > 0 Then

102         For i = 1 To MAX_INVENTORY_SLOTS

104             If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
106                 QuedanItems = True

                    Exit Function

                End If

            Next

        End If

108     QuedanItems = False
        '<EhFooter>
        Exit Function

QuedanItems_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.QuedanItems " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo EncontrarCant_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 03/09/08
        'Last Modification By: Marco Vanotti (Marco)
        ' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
        '***************************************************

        'Devuelve la cantidad original del obj de un npc

        Dim ln As String, npcfile As String

        Dim i  As Integer
    
100     npcfile = Npcs_FilePath
     
102     For i = 1 To MAX_INVENTORY_SLOTS
104         ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).numero, "Obj" & i)

106         If ObjIndex = val(ReadField(1, ln, 45)) Then
108             EncontrarCant = val(ReadField(2, ln, 45))

                Exit Function

            End If

        Next
                       
110     EncontrarCant = 0

        '<EhFooter>
        Exit Function

EncontrarCant_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.EncontrarCant " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo ResetNpcInv_Err
        '</EhHeader>


        Dim i As Integer
    
100     With Npclist(NpcIndex)
102         .Invent.NroItems = 0
104         .NroDrops = 0
        
106         For i = 1 To MAX_INVENTORY_SLOTS
108             .Invent.Object(i).ObjIndex = 0
110             .Invent.Object(i).Amount = 0
112             .Drop(i).ObjIndex = 0
114             .Drop(i).Amount = 0
116             .Drop(i).Probability = 0
                  .Drop(i).ProbNum = 0
118         Next i
        
120         .InvReSpawn = 0
        End With

        '<EhFooter>
        Exit Sub

ResetNpcInv_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.ResetNpcInv " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, _
                     ByVal Slot As Byte, _
                     ByVal cantidad As Integer)
        '<EhHeader>
        On Error GoTo QuitarNpcInvItem_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 23/11/2009
        'Last Modification By: Marco Vanotti (Marco)
        ' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        Dim ObjIndex As Integer

        Dim iCant    As Integer
    
100     With Npclist(NpcIndex)
102         ObjIndex = .Invent.Object(Slot).ObjIndex
    
            'Quita un Obj
104         If ObjData(.Invent.Object(Slot).ObjIndex).Crucial = 0 Then
106             .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - cantidad
            
108             If .Invent.Object(Slot).Amount <= 0 Then
110                 .Invent.NroItems = .Invent.NroItems - 1
112                 .Invent.Object(Slot).ObjIndex = 0
114                 .Invent.Object(Slot).Amount = 0
                
116                 If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
118                     Call CargarInvent(NpcIndex) 'Reponemos el inventario
                    End If
                
                End If

            Else
120             .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - cantidad
            
122             If .Invent.Object(Slot).Amount <= 0 Then
124                 .Invent.NroItems = .Invent.NroItems - 1
126                 .Invent.Object(Slot).ObjIndex = 0
128                 .Invent.Object(Slot).Amount = 0
                
130                 If Not QuedanItems(NpcIndex, ObjIndex) Then
                        'Check if the item is in the npc's dat.
132                     iCant = EncontrarCant(NpcIndex, ObjIndex)

134                     If iCant Then
136                         .Invent.Object(Slot).ObjIndex = ObjIndex
138                         .Invent.Object(Slot).Amount = iCant
140                         .Invent.NroItems = .Invent.NroItems + 1
                        End If
                    End If
                
142                 If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
144                     Call CargarInvent(NpcIndex) 'Reponemos el inventario
                    End If
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

QuitarNpcInvItem_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.QuitarNpcInvItem " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        '<EhHeader>
        On Error GoTo CargarInvent_Err
        '</EhHeader>

        'Vuelve a cargar el inventario del npc NpcIndex
        Dim LoopC   As Integer

        Dim ln      As String

        Dim npcfile As String
    
100     npcfile = Npcs_FilePath
    
102     With Npclist(NpcIndex)
104         .Invent.NroItems = val(GetVar(npcfile, "NPC" & .numero, "NROITEMS"))
        
106         For LoopC = 1 To .Invent.NroItems
108             ln = GetVar(npcfile, "NPC" & .numero, "Obj" & LoopC)
110             .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
112             .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            
114         Next LoopC

        End With

        '<EhFooter>
        Exit Sub

CargarInvent_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.CargarInvent " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub TirarOroNpc(ByVal UserIndex As Integer, _
                       ByVal cantidad As Long, _
                       ByRef Pos As WorldPos)
        '<EhHeader>
        On Error GoTo TirarOroNpc_Err
        '</EhHeader>

        '***************************************************
        'Autor: ZaMa
        'Last Modification: 13/02/2010
        '***************************************************

100     If cantidad > 0 And UserIndex > 0 Then
        
102         With UserList(UserIndex)


104                 If UserList(UserIndex).Stats.BonusTipe = eEffectObj.e_Gld Then

                        Dim Diferencia As Long
                    
106                     Diferencia = cantidad * UserList(UserIndex).Stats.BonusValue
108                     Diferencia = Diferencia - cantidad

                    End If
            
110                 .Stats.Gld = .Stats.Gld + cantidad + Diferencia
                
112                 Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("Oro +" & CStr(Format(cantidad + Diferencia, "###,###,###")), d_AddGld, 3000, 0))
114                 WriteUpdateGold (UserIndex)
            
            End With
        
        End If


        '<EhFooter>
        Exit Sub

TirarOroNpc_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.InvNpc.TirarOroNpc " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

