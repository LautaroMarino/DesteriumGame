Attribute VB_Name = "Protocol_Out"



' Todos los paquetes que el protocolo envía al cliente

Option Explicit


Public Sub WriteMultiMessage(ByVal UserIndex As Integer, _
                             ByVal MessageIndex As Integer, _
                             Optional ByVal Arg1 As Long, _
                             Optional ByVal Arg2 As Long, _
                             Optional ByVal Arg3 As Long, _
                             Optional ByVal StringArg1 As String)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteInt(ServerPacketID.MultiMessage)
        Call .WriteInt(MessageIndex)
        
        Select Case MessageIndex

            Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
                eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.SafeCvcModeOn, eMessages.SafeCvcModeOff, eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteInt(Arg1) 'Target
                Call .WriteInt(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call .WriteInt(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call .WriteInt(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInt(Arg1) 'AttackerIndex
                Call .WriteInt(Arg2) 'Target
                Call .WriteInt(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call .WriteInt(Arg1) 'AttackerIndex
                Call .WriteInt(Arg2) 'Target
                Call .WriteInt(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call .WriteInt(Arg1) 'skill
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
                Call .WriteInt(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call .WriteInt(Arg2) 'Expe
            
            Case eMessages.UserKill '"¡" & .name & " te ha matado!"
                Call .WriteInt(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call .WriteInt(CByte(Arg1))
                Call .WriteInt(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda así.
                Call .WriteASCIIString(StringArg1) 'Call .WriteInt(CByte(Arg2))
                
        End Select

    End With

    Exit Sub ''

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)

        Resume

    End If

End Sub

