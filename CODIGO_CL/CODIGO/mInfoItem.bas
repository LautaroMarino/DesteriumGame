Attribute VB_Name = "mInfoItem"
Option Explicit

Public LastDelay As Long

Public Sub ShowInfoItem(ByVal ObjIndex As Integer)
    
    Dim L     As Long

    Dim T     As Long

    Dim Delay As Long
        
    If SelectedObjIndex <> ObjIndex Then
        If (FrameTime - LastDelay) <= 100 Then
            Exit Sub

        End If
        
        LastDelay = FrameTime

        SelectedObjIndex = ObjIndex
        
        If Not MirandoObjetos Then
            If MirandoComerciar Then
                FrmObject_Info.Show ', frmComerciar
            ElseIf MirandoBanco Then
                FrmObject_Info.Show ', frmBancoObj
            Else

                FrmObject_Info.Show

            End If

        Else
    
            FrmObject_Info.Initial_Form

        End If
        
    Else

        If MirandoObjetos Then
            If Not FrmObject_Info.visible Then
                FrmObject_Info.visible = True
                FrmObject_Info.Initial_Form

            End If

        Else
          
            FrmObject_Info.Initial_Form

        End If

    End If

    If MirandoComerciar Then
        If frmComerciar.SolapaView = 1 Then
            L = frmComerciar.Left + (frmComerciar.picInvNpc.Left * Screen.TwipsPerPixelX) + (frmComerciar.MouseX * Screen.TwipsPerPixelX)
            T = frmComerciar.Top + (frmComerciar.picInvNpc.Top * Screen.TwipsPerPixelX) + (frmComerciar.MouseY * Screen.TwipsPerPixelY) + (32 * Screen.TwipsPerPixelY)
        ElseIf frmComerciar.SolapaView = 2 Then
            L = frmComerciar.Left + (frmComerciar.picInvUser.Left * Screen.TwipsPerPixelX) + (frmComerciar.MouseX * Screen.TwipsPerPixelX)
            T = frmComerciar.Top + (frmComerciar.picInvUser.Top * Screen.TwipsPerPixelX) + (frmComerciar.MouseY * Screen.TwipsPerPixelY) + (32 * Screen.TwipsPerPixelY)

        End If
        
        L = frmComerciar.Left + frmComerciar.Width + 20
        T = frmComerciar.Top
        
    ElseIf MirandoBanco Then

        If frmBancoObj.SolapaView = 1 Then
            L = frmBancoObj.Left + (frmBancoObj.PicBancoInv * Screen.TwipsPerPixelX) + (frmBancoObj.MouseX * Screen.TwipsPerPixelX)
            T = frmBancoObj.Top + (frmBancoObj.PicBancoInv.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (frmBancoObj.MouseY * Screen.TwipsPerPixelY)
        ElseIf frmBancoObj.SolapaView = 2 Then
            L = frmBancoObj.Left + (frmBancoObj.PicInv.Left * Screen.TwipsPerPixelX) + (frmBancoObj.MouseX * Screen.TwipsPerPixelX)
            T = frmBancoObj.Top + (frmBancoObj.PicInv.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (frmBancoObj.MouseY * Screen.TwipsPerPixelY)

        End If
        
        L = frmBancoObj.Left + frmBancoObj.Width + 20
        T = frmBancoObj.Top
        
    ElseIf MirandoListaDrops Then
        L = FrmMapa.Left + (FrmMapa.PicMapa.Left * Screen.TwipsPerPixelX) + (FrmMapa.MouseX * Screen.TwipsPerPixelX)
        T = FrmMapa.Top + (FrmMapa.PicMapa.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (FrmMapa.MouseY * Screen.TwipsPerPixelY)

    ElseIf MirandoListaCofres Then
        'L = FrmMapa.Left + (FrmMapa.picCofreItem.Left * Screen.TwipsPerPixelX) + (FrmMapa.MouseX * Screen.TwipsPerPixelX)
        'T = FrmMapa.Top + (FrmMapa.picCofreItem.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (FrmMapa.MouseY * Screen.TwipsPerPixelY)
    ElseIf MirandoSkins Then
        L = FrmSkin.Left + (FrmSkin.PicInv.Left * Screen.TwipsPerPixelX) + (FrmSkin.MouseX * Screen.TwipsPerPixelX)
        T = FrmSkin.Top + (FrmSkin.PicInv.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (FrmSkin.MouseY * Screen.TwipsPerPixelY)
        
        L = FrmSkin.Left + FrmSkin.Width + 20
        T = FrmSkin.Top
    Else ' Ventana principal
        L = FrmMain.Left + (FrmMain.PicInv.Left * Screen.TwipsPerPixelX) + (FrmMain.MouseX * Screen.TwipsPerPixelX)
        T = FrmMain.Top + (FrmMain.PicInv.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (FrmMain.MouseY * Screen.TwipsPerPixelY)

    End If

    ' Update View in border
    #If ModoBig > 0 Then

        If (L + 2500) > FrmMain.Width Then L = L - 2500
        If (T + 2500) > FrmMain.Height Then T = T - 2500
    #Else

        If ClientSetup.bConfig(eSetupMods.SETUP_PANTALLACOMPLETA) = 1 Then
            If (L + 2200) > FrmMain.Width Then L = L - 2200
            If (T + 2200) > FrmMain.Height Then T = T - 2200

        End If

    #End If
    
    FrmObject_Info.Move L, T

End Sub


