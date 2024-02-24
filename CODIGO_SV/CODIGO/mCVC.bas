Attribute VB_Name = "mCVC"
' Sistema de CVC
' Dragons AO 2020
' Je

Option Explicit

Public Const MIN_USERS_CVC As Byte = 3

Public Type tCVCMap

    Occuped As Boolean
    Map(1) As Integer
    X(1) As Byte
    Y(1) As Byte

End Type

Public Type tCVC

    Run As Boolean
    User() As String
    MapSelected As Byte
    max As Byte
    
End Type

Public CVC(0)      As tCVC

Public CVC_MAPS(0) As tCVCMap

'#################################################
'#################################################
' PROCEDIMIENTOS PROPIOS DEL CVC

Public Sub CVC_LoadConfig()

    With CVC_MAPS(0)
        .Occuped = False
        
        .Map(0) = 286
        .X(0) = 50
        .Y(0) = 16
        
        .Map(1) = 286
        .X(1) = 50
        .Y(1) = 85
    End With
    
End Sub

Private Sub CVC_SetNew(ByVal max As Byte, ByRef Users() As String)
    
    With CVC(0)
        .Run = True
        .User = Users
        
        Call CVC_SetFlagsUsers(Users)
    End With
    
End Sub

Private Sub CVC_Reset()

    Dim tUser As Integer

    Dim A     As Long
    
    With CVC(0)
        .Run = False
        .MapSelected = 0
        
        For A = LBound(.User) To UBound(.User)
            tUser = .User(A)
            
            If tUser > 0 Then
                Call CVC_ResetUser(tUser)
            End If

        Next A
        
    End With
    
End Sub

Public Sub CVC_ResetUser(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .CvcTemp.LastRequest = 0
        .CvcTemp.Team = 0
    End With
    
End Sub

' Seteamos los flags en los personajes involucrados.
Private Sub CVC_SetFlagsUsers(ByRef Users() As String)

    On Error GoTo Errhandler

    Dim A     As Long

    Dim tUser As Integer
    
    For A = LBound(Users) To UBound(Users)
        tUser = NameIndex(Users(A))
        
        If A <= (UBound(Users) / 2) Then
            UserList(tUser).CvcTemp.Team = 1
            Call EventosDS.EventWarpUser(tUser, CVC_MAPS(0).Map(0), CVC_MAPS(0).X(0), CVC_MAPS(0).Y(0))
        Else
            UserList(tUser).CvcTemp.Team = 2
            Call EventosDS.EventWarpUser(tUser, CVC_MAPS(0).Map(1), CVC_MAPS(0).X(1), CVC_MAPS(0).Y(1))
        End If
        
        LogError ("TESTEO» " & Users(A) & " es del TeamCVC=" & UserList(tUser).CvcTemp.Team)
    Next A
    
    Exit Sub

Errhandler:
    Call LogError("TESTEOCVC» ERROR EN CVC_SetFlagsUsers")
End Sub

'#################################################
'#################################################
' PROCEDIMIENTOS QUE HACE EL USUARIO
' Un personaje LIDER envía solicitud a otro personaje LIDER

Public Function CVC_CanUserFight(ByVal UserIndex As Integer, _
                                 Optional ByVal CheckGuild As Boolean = False) As Boolean

    With UserList(UserIndex)
        
        If CheckGuild Then
            If .GuildIndex = 0 Then Exit Function
            'If UCase$(GuildLeader(.GuildIndex)) <> UCase$(.Name) Then Exit Function
        End If
        
        If .Counters.Pena Then Exit Function
        If .Counters.Saliendo Then Exit Function
        If .flags.SlotEvent > 0 Then Exit Function
        If .flags.SlotReto > 0 Then Exit Function
        If .flags.Desafiando > 0 Then Exit Function
        If .flags.Envenenado > 0 Then Exit Function
        
        CVC_CanUserFight = True
    End With

End Function

Public Sub CVC_SendInvitation(ByVal UserIndex As Integer, ByVal UserLeader As Integer)

    On Error GoTo Errhandler
    
    With UserList(UserIndex)
        
        If .CvcTemp.LastRequest = UserList(UserLeader).GuildIndex Then
            Call WriteConsoleMsg(UserIndex, "Ya has enviado una invitación. Espera unos momentos...", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If CVC_CanUserFight(UserIndex, True) Then
            Call WriteConsoleMsg(UserIndex, "Tu personaje no puede enviar solicitud de CVC en estos momentos.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If CVC_CanUserFight(UserLeader, True) Then
            Call WriteConsoleMsg(UserIndex, "El personaje " & UserList(UserLeader).Name & " no puede recibir solicitud de CVC en estos momentos.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        .CvcTemp.LastRequest = UserList(UserLeader).GuildIndex
        
        'Call WriteConsoleMsg(UserLeader, "El lider del clan " & GuildName(.GuildIndex) & " te ha enviado una solicitud a enfrentamiento Clan vs Clan. Para aceptar tipea /SICVC " & .Name, FontTypeNames.FONTTYPE_GUILD)
    End With
    
    Exit Sub

Errhandler:
    Call LogError("TESTEOCVC» ERROR EN CVC_SendInvitation")
End Sub

' Un personaje LIDER acepta la solicitud recibida por otro personaje LIDER
Public Sub CVC_AcceptInvitation(ByVal UserIndex As Integer, ByVal UserLeader As Integer)

    Exit Sub

Errhandler:
    Call LogError("TESTEOCVC» ERROR EN CVC_AcceptInvitation")
End Sub

' Buscamos un posible Atacante. Ya que al deslogear, no existe atacante, buscamos al team opuesto.
Private Function CVC_SearchAttackerIndex(ByVal VictimIndex As Integer) As Integer

    On Error GoTo Errhandler
    
    Dim A     As Long

    Dim tUser As String
    
    For A = LBound(CVC(0).User) To UBound(CVC(0).User)
        tUser = NameIndex(CVC(0).User(A))
        
        If tUser > 0 Then
            If (UserList(tUser).CvcTemp.Team > 0) And (UserList(tUser).CvcTemp.Team <> UserList(VictimIndex).CvcTemp.Team) Then
                
                CVC_SearchAttackerIndex = tUser

                Exit Function

            End If
        End If

    Next A
    
    Exit Function

Errhandler:
    Call LogError("TESTEOCVC» ERROR EN CVC_SearchAttackerIndex")
End Function

' Un personaje muere en el CVC..
Public Sub CVC_Userdie(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

    On Error GoTo Errhandler

    'Si el CVC puede continuar salimos de acá
    If CVC_CheckContinue(VictimIndex) Then Exit Sub
    
    If AttackerIndex = 0 Then
        AttackerIndex = CVC_SearchAttackerIndex(VictimIndex)
    End If
    
    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("CVC» El clan " & GuildName(UserList(AttackerIndex).GuildIndex) & " ha derrotado al clan " & GuildName(UserList(VictimIndex).GuildIndex) & ". ¡Ha ganado 50 GuildPoints!", FontTypeNames.FONTTYPE_GUILD))
    Call CVC_Reset
    
    Exit Sub

Errhandler:
    Call LogError("TESTEOCVC» ERROR EN CVC_USERDIE")
End Sub

' Buscamos al menos un miembro del clan víctima que esté vivo para continuar..
Private Function CVC_CheckContinue(ByVal VictimIndex As Integer) As Boolean

    On Error GoTo Errhandler

    Dim A     As Long

    Dim tUser As String
    
    For A = LBound(CVC(0).User) To UBound(CVC(0).User)
        tUser = NameIndex(CVC(0).User(A))
        
        If tUser > 0 Then
            If (UserList(tUser).CvcTemp.Team > 0) And (UserList(tUser).CvcTemp.Team = UserList(VictimIndex).CvcTemp.Team) And (UserList(tUser).flags.Muerto = 0) Then
                
                CVC_CheckContinue = True

                Exit Function

            End If
        End If

    Next A
    
    Exit Function

Errhandler:
    Call LogError("TESTEOCVC» ERROR EN CVC_CheckContinue")
End Function
