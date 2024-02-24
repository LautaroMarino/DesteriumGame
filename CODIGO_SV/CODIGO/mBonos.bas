Attribute VB_Name = "mBonos"
' Creado por DragonsAO

Option Explicit

' Function
Public Function IsUserBono(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).flags.SelectedBono > 0 Then IsUserBono = True
End Function

Public Sub Bonos_UserFinish(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        .flags.SelectedBono = 0
        .Counters.TimeBono = 0
        
        WriteConsoleMsg UserIndex, "El tiempo del bono ha terminado. Ya no sentirás el efecto.", FontTypeNames.FONTTYPE_INFO
    End With

End Sub

