Attribute VB_Name = "modMenues"
Option Explicit

Public DisplayingMenu As Byte

Public Sub LoadMenuInfo()

    On Error Resume Next

    Dim FilePath As String
    
    FilePath = IniPath & "Menu.dat"
    
    If Not FileExist(FilePath, vbArchive) Then
        Call MsgBox("ERROR: no se ha podido cargar los menues. Falta el archivo menu.dat, reinstale el juego", vbCritical + vbOKOnly)

        Exit Sub

    End If
    
    Dim NumMenues As Long

    NumMenues = Val(GetVar(FilePath, "INIT", "NumMenues"))
    
    If NumMenues > 0 Then ReDim MenuInfo(1 To NumMenues)
    
    Dim Index       As Long

    Dim ActionIndex As Long

    Dim iTemp       As Integer
    
    For Index = 1 To NumMenues
        
        iTemp = Val(GetVar(FilePath, "MENU" & Index, "NumActions"))
        
        With MenuInfo(Index)
            .NumActions = CByte(iTemp)
        
            If iTemp > 0 Then
                ReDim .Actions(1 To iTemp)
                
                For ActionIndex = 1 To iTemp
                    .Actions(ActionIndex).ActionIndex = Val(GetVar(FilePath, "MENU" & Index, "Action" & ActionIndex))
                    .Actions(ActionIndex).NormalGrh = Val(GetVar(FilePath, "MENU" & Index, "NormalGrh" & ActionIndex))
                    .Actions(ActionIndex).FocusGrh = Val(GetVar(FilePath, "MENU" & Index, "FocusGrh" & ActionIndex))
                Next ActionIndex

            End If

        End With

    Next Index
    
End Sub

Public Sub PerformMenuAction(ByVal Action As Byte)

    Debug.Print "Perform: " & Action

    Select Case Action
    
        Case eMenuAction.ieCommerce
            Call WriteCommerceStart
            
        Case eMenuAction.iePriestHeal
            Call WriteHeal
            
        Case eMenuAction.ieHogar
            'Call WriteHome

        Case eMenuAction.ieBank
            Call WriteBankStart(E_BANK.e_User)
            
        Case eMenuAction.ieFactionEnlist
            ' Call WriteEnlist
            
        Case eMenuAction.ieFactionReward
            'Call WriteReward
            
        Case eMenuAction.ieFactionWithdraw
            'Call WriteLeaveFaction
            
        Case eMenuAction.ieFactionInfo
            '  Call WriteInformation
            
        Case eMenuAction.ieGamble
            '*
        
        Case eMenuAction.ieBlacksmith
            '*
        
        Case eMenuAction.ieMakeLingot
            '*
        
        Case eMenuAction.ieMeltDown
            '*
        
        Case eMenuAction.ieShareNpc
            Call WriteShareNpc
            
        Case eMenuAction.ieStopSharingNpc
            Call WriteStopSharingNpc
            
        Case eMenuAction.ieTameNpc
            '*
        
        Case eMenuAction.ieMakeFireWood
            '*
        
        Case eMenuAction.ieLightFire
            '*
    End Select

End Sub
