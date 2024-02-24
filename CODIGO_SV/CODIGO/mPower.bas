Attribute VB_Name = "mPower"
' Programado para Desterium AO EXODO III
Option Explicit

Private Type tPower
    Desc As String
    UserIndex As Integer
    FindMap As Boolean
    PreviousUser As String
    Active As Boolean
    Time As Integer
End Type

Private NumPower As Byte

Public Power  As tPower

Public Sub Power_Search(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler

    Dim A         As Long, B As Long, C As Long
    
    If (NumUsers + UsersBot) < 30 Then Exit Sub
    If Power.UserIndex > 0 Then Exit Sub
    
    With Power
        
        For A = 2 To NumMaps
            With MapInfo(A)
                If .Poder = 1 Then
                    If Power.UserIndex = 0 Then
                        If Not StrComp(UCase$(Power.PreviousUser), UCase$(UserList(UserIndex).Name)) = 0 Then
                            If UserList(UserIndex).flags.UserLogged And _
                                UserList(UserIndex).flags.Muerto = 0 And Not EsGm(UserIndex) And _
                                UserList(UserIndex).Clase <> eClass.Thief Then
                                    
                                Call Power_Set(UserIndex, Power.UserIndex)
                                Call Power_Message
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_WARP, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 52, 5))
                                'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(128, NO_3D_SOUND, NO_3D_SOUND))
                                
                                Exit Sub
                                    
                            End If
                        End If
                    End If
                End If
            End With

        Next A
    
    End With

Exit Sub
ErrHandler:
End Sub

Public Sub Power_Search_All()
    On Error GoTo ErrHandler
    
    'If Not Power.Active Then Exit Sub
    If (NumUsers + UsersBot) < 30 Then Exit Sub
    
    Dim A         As Long
    
    For A = 1 To LastUser
        If UserList(A).flags.UserLogged Then
            Call Power_Search(A)
        End If
    Next A

Exit Sub
ErrHandler:
End Sub

Public Sub Power_Set(ByVal UserIndex As Integer, _
                      ByVal PreviousUser As Integer)
        '<EhHeader>
        On Error GoTo Power_Set_Err
        '</EhHeader>
    
100     With Power
102         If Not PreviousUser = 0 Then
104             .PreviousUser = UCase$(UserList(PreviousUser).Name)
            End If
            
106         .UserIndex = UserIndex
108         .Time = 900 '1800
        
            ' Si el UserIndex se resetea, buscamos al nuevo poder
110         If .UserIndex = 0 Then
            
112             Call Power_Search_All
            Else
114             Call RefreshCharStatus(.UserIndex)
            End If
        End With

        '<EhFooter>
        Exit Sub

Power_Set_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mPower.Power_Set " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Power_Message()
    
    On Error GoTo ErrHandler
    
    With Power
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Desc & UserList(.UserIndex).Name, FontTypeNames.FONTTYPE_GUILD))
    End With
    
    Exit Sub
ErrHandler:
    
End Sub

Public Function Power_CheckTime() As Boolean
        '<EhHeader>
        On Error GoTo Power_CheckTime_Err
        '</EhHeader>
        Dim Time As String
    
        'Time = Format(Now, "hh:mm")
    
100     If Power.Time > 0 Then
102         Power.Time = Power.Time - 1
        
104         If Power.Time = 0 Then
106             Call WriteConsoleMsg(Power.UserIndex, "Has perdido el poder", FontTypeNames.FONTTYPE_INFORED)
108             Call RefreshCharStatus(Power.UserIndex)
110             Call Power_Set(0, Power.UserIndex)
            End If
        End If
    
 
        '<EhFooter>
        Exit Function

Power_CheckTime_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mPower.Power_CheckTime " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
