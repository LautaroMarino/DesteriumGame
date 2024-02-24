Attribute VB_Name = "mAutoBan"
Option Explicit

Private Type tAutoBan

    Name As String
    cant As Long
    Time As Long
    Reason As String

End Type

Public AutoBan()                    As tAutoBan

Public LastAutoBan                  As Integer

Private Const AUTOBAN_MAX_TOLERANCE As Byte = 10
Private Const AUTOBAN_TIME          As Long = 14400 ' 4 HORAS

Public Sub AutoBan_Initialize()
        '<EhHeader>
        On Error GoTo AutoBan_Initialize_Err
        '</EhHeader>
    
100     ReDim AutoBan(0) As tAutoBan
    
        '<EhFooter>
        Exit Sub

AutoBan_Initialize_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAutoBan.AutoBan_Initialize " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AutoBan_AddUser(ByVal UserName As String, _
   ByVal Reason As String)
                            
    On Error GoTo ErrHandler
    
    Dim Slot As Integer
    
    Slot = AutoBan_Repeat(UserName)
    
    If Slot > 0 Then

        With AutoBan(Slot)
            .cant = .cant + 1
                
            If .cant = AUTOBAN_MAX_TOLERANCE Then
                .Time = AUTOBAN_TIME

                Exit Sub

            End If
                
        End With
            
    Else
        Slot = AutoBan_SlotFree
            
        If Slot = 0 Then
            ReDim Preserve AutoBan(LBound(AutoBan) To UBound(AutoBan) + 1) As tAutoBan
            Slot = UBound(AutoBan)
        End If
        
        With AutoBan(Slot)
            .Name = UserName
            .cant = 1
            .Time = 0
            .Reason = Reason
        End With

    End If
    
    Exit Sub

ErrHandler:
    Call LogError("Error en AutoBan_AddUser")
    
End Sub

Public Sub AutoBan_RemoveUser(ByVal Slot As Long)
        '<EhHeader>
        On Error GoTo AutoBan_RemoveUser_Err
        '</EhHeader>

100     With AutoBan(Slot)
102         .Name = vbNullString
104         .cant = 0
106         .Time = 0
108         .Reason = vbNullString
        End With

        '<EhFooter>
        Exit Sub

AutoBan_RemoveUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAutoBan.AutoBan_RemoveUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AutoBan_Character(ByVal UserName As String, _
   ByVal Reason As String)
        '<EhHeader>
        On Error GoTo AutoBan_Character_Err
        '</EhHeader>

        Dim tUser As Integer

        Dim Penas As Integer
    
100     tUser = NameIndex(UserName)
    
102     Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
    
104     Penas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
106     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Penas + 1)
108     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Penas + 1, ": BAN POR Macro Externo " & Date & " " & Time)
    
110     If tUser > 0 Then
112         UserList(tUser).flags.Ban = 1
            'Call FlushBuffer(tUser)
            'Call CloseSocket(tUser)
        
114         Call WriteDisconnect(tUser)
116         Call FlushBuffer(tUser)
                        
118         Call CloseSocket(tUser)
        End If
    
120     Call Logs_Security(eSecurity, eAutoBan, "Personaje " & UserName & " BAN por AntiCheat automático. Razon Real: " & Reason)
    
        '<EhFooter>
        Exit Sub

AutoBan_Character_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAutoBan.AutoBan_Character " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AutoBan_Loop()
        '<EhHeader>
        On Error GoTo AutoBan_Loop_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = LBound(AutoBan) To UBound(AutoBan)

102         With AutoBan(A)

104             If .Time > 0 Then
106                 .Time = .Time - 1
                
108                 If .Time = 0 Then
110                     If GetVar(CharPath & .Name & ".chr", "FLAGS", "Ban") = "0" Then
112                         Call AutoBan_Character(.Name, .Reason)
                        End If
                    
114                     Call AutoBan_RemoveUser(A)
                    End If
                End If
        
            End With
    
116     Next A

        '<EhFooter>
        Exit Sub

AutoBan_Loop_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAutoBan.AutoBan_Loop " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' FUNCIONES
Private Function AutoBan_Repeat(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo AutoBan_Repeat_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = LBound(AutoBan) To UBound(AutoBan)

102         With AutoBan(A)

104             If StrComp(UserName, .Name) = 0 Then
106                 AutoBan_Repeat = A

                    Exit Function

                End If

            End With

108     Next A
    
        '<EhFooter>
        Exit Function

AutoBan_Repeat_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAutoBan.AutoBan_Repeat " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Private Function AutoBan_SlotFree()
        '<EhHeader>
        On Error GoTo AutoBan_SlotFree_Err
        '</EhHeader>

        Dim A As Long
    
100     For A = 1 To UBound(AutoBan)

102         If AutoBan(A).Name = vbNullString Then
104             AutoBan_SlotFree = A

                Exit Function

            End If

106     Next A
    
        '<EhFooter>
        Exit Function

AutoBan_SlotFree_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAutoBan.AutoBan_SlotFree " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

