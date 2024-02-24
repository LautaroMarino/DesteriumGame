Attribute VB_Name = "mBattle"
Option Explicit

#If Classic = 0 Then
' Se viene la mejor revolucion del AO


Public ARENA_LAST As Long

Public Type tBattleArena
        Name As String
        Maps() As Integer
        Limit As Byte
        
        
        ' Reseteable
        Users As Byte
        
End Type

Public Battle_Arenas() As tBattleArena

Public Sub IBatle_LoadArenas()
        '<EhHeader>
        On Error GoTo IBatle_LoadArenas_Err
        '</EhHeader>
    
        Dim Read As clsIniManager
        Dim FilePath As String
        Dim A As Long, B As Long
        Dim Maps As String
        Dim List() As String
100     Set Read = New clsIniManager
    
102     FilePath = DatPath & "ARENAS.ini"
104     Read.Initialize FilePath
    
106     ARENA_LAST = val(Read.GetValue("INIT", "LAST"))
    
108     ReDim Battle_Arenas(1 To ARENA_LAST) As tBattleArena
    
110     For A = 1 To ARENA_LAST
112         With Battle_Arenas(A)
114             .Name = Read.GetValue(A, "NAME")
116             .Limit = val(Read.GetValue(A, "LIMIT"))
118             Maps = Read.GetValue(A, "MAPS")
120             List = Split(Maps, "-")
            
122             ReDim .Maps(LBound(List) To UBound(List)) As Integer
124             For B = LBound(List) To UBound(List)
126                 .Maps(B) = val(List(B))
                
128             Next B
            End With
    
130     Next A
    
132     Set Read = Nothing
        '<EhFooter>
        Exit Sub

IBatle_LoadArenas_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mBattle.IBatle_LoadArenas " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

#End If
