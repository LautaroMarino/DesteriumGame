Attribute VB_Name = "mSkills"
Option Explicit


Public Sub Passive_Add(ByVal UserIndex As Integer, ByVal Value As Long)

    With UserList(UserIndex)
        .CharacterStats.PassiveAccumulated = .CharacterStats.PassiveAccumulated + Value
        
        
    
    End With
    
End Sub
