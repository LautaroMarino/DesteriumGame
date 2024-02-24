Attribute VB_Name = "mMeditations"
Option Explicit

Public Const MAX_MEDITATION            As Byte = 39
Public Const MEDITATION_NPC            As Integer = 18
Public Meditation(0 To MAX_MEDITATION) As Integer

Public Sub Meditation_LoadConfig()
    
    
    Meditation(0) = 0
    Meditation(1) = 105
    Meditation(2) = 106
    Meditation(3) = 107
    Meditation(4) = 108
    
    Meditation(5) = 109
    Meditation(6) = 110
    Meditation(7) = 111
    Meditation(8) = 112
    
    Meditation(9) = 113
    Meditation(10) = 114
    Meditation(11) = 115
    Meditation(12) = 116
    
    Meditation(13) = 117
    Meditation(14) = 118
    Meditation(15) = 119
    Meditation(16) = 120
    
    ' Las del Tigre
    Meditation(17) = 121
    Meditation(18) = 122
    Meditation(19) = 123
    Meditation(20) = 124
    
    
    ' Las que son horrendas pero les gusta
    Meditation(21) = 31
    Meditation(22) = 32
    Meditation(23) = 33
    Meditation(24) = 34
    Meditation(25) = 35
    Meditation(26) = 36
    Meditation(27) = 37
    Meditation(28) = 38
    
    ' Las de la nube que son mas feas todavia
    Meditation(29) = 39
    Meditation(30) = 40
    Meditation(31) = 41
    
    ' Las de Los rayos que se ven como el orto pero les gusta
    Meditation(32) = 62
    Meditation(33) = 63
    Meditation(34) = 64
    Meditation(35) = 65
    Meditation(36) = 66
    Meditation(37) = 67
    Meditation(38) = 68
    Meditation(39) = 69
End Sub

Public Sub Meditation_Select(ByVal UserIndex As Integer, ByVal Selected As Byte)

        '<EhHeader>
        On Error GoTo Meditation_Select_Err

        '</EhHeader>
    
100     With UserList(UserIndex)

102         If .MeditationSelected = Selected Then
104             Call WriteErrorMsg(UserIndex, "Ya tienes elegida esa concentración.")

                Exit Sub

            End If
             
106
        
108         If Selected Then
                If .MeditationUser(Selected) = 0 Then Exit Sub  ' Como llego hasta aca

                .MeditationSelected = Selected
110           Call WriteErrorMsg(UserIndex, "Has elegido usar la meditación n° " & Selected)
                
            Else
                .MeditationSelected = 0
112             Call WriteErrorMsg(UserIndex, "Has elegido usar la meditación por defecto.")

            End If
            
            Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageUpdateMeditation(UserList(UserIndex).MeditationUser, Meditation(Selected)))
        End With
    
        '<EhFooter>
        Exit Sub

Meditation_Select_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMeditations.Meditation_Select " & "at line " & Erl

        

        '</EhFooter>
End Sub
