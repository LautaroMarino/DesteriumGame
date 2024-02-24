Attribute VB_Name = "mFilePath"
Option Explicit
' Desterium presentará varios MODS
' Conversión del Path según versión del Juego


Public Function Npcs_FilePath() As String
    #If Classic = 1 Then
            Npcs_FilePath = DatPath & "Npcs.dat"
    #Else
            Npcs_FilePath = DatPath & "NpcsClassic.dat"
    #End If
End Function

Public Function Objs_FilePath() As String
    #If Classic = 1 Then
            Objs_FilePath = DatPath & "Obj.dat"
    #Else
            Objs_FilePath = DatPath & "ObjClassic.dat"
    #End If
End Function

Public Function Spell_FilePath() As String
    #If Classic = 1 Then
            Spell_FilePath = DatPath & "Hechizos.dat"
    #Else
            Spell_FilePath = DatPath & "HechizosClassic.dat"
    #End If
End Function

Public Function Maps_FilePath() As String
    #If Classic = 1 Then
            Maps_FilePath = App.Path & "\MAPS\"
    #Else
            Maps_FilePath = App.Path & "\MAPSCLASSIC\"
    #End If
End Function

Public Function Quests_FilePath() As String
    #If Classic = 1 Then
            Quests_FilePath = DatPath & "Quest.dat"
    #Else
            Quests_FilePath = DatPath & "QuestClassic.dat"
    #End If
End Function
Public Function Drops_FilePath() As String
    #If Classic = 1 Then
            Drops_FilePath = DatPath & "DROP.DAT"
    #Else
            Drops_FilePath = DatPath & "DropsC.dat"
    #End If
End Function
Public Function Drops_FilePath_Client() As String
    #If Classic = 1 Then
            Drops_FilePath_Client = DatPath & "client\server_drops.ind"
    #Else
            Drops_FilePath_Client = DatPath & "client\server_drops_classic.ind"
    #End If
End Function
