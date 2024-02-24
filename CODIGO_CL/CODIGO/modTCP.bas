Attribute VB_Name = "Mod_TCP"
Option Explicit


Public Sub Prepare_And_Connect(ByVal Mode As E_MODO, Optional ByRef Parent As Form = Nothing, Optional ByRef sIPAddress As String = vbNullString, Optional ByVal nPort As Integer = 0)
  
    ' @ Prepara y conecta el socket.
     
    EstadoLogin = Mode
    
    If (modNetwork.IsConnected) Then
        Call Login
    Else
        If LenB(sIPAddress) > 0 Then
            Call modNetwork.Connect(sIPAddress, nPort)
        Else
            Call modNetwork.Connect(CurServerIp, CurServerPort)
        End If
    End If

End Sub

Public Sub Login()
    Select Case EstadoLogin
        Case E_MODO.e_LoginMercaderOff
            Call WriteMercader_Required(MercaderOff, 0, 0)
            
        Case E_MODO.e_LoginAccountPasswd
            Call WriteLoginPasswd
            
        Case E_MODO.e_LoginAccountChar
            Call WriteLoginChar
        
        Case E_MODO.e_LoginAccountNewChar
            Call WriteLoginCharNew
            
        Case E_MODO.e_LoginName
            Call WriteLoginName
            
        Case E_MODO.e_LoginAccount
            Call WriteLoginAccount
            
        Case E_MODO.e_LoginAccountNew
            Call WriteLoginAccountNew
            
        Case E_MODO.e_LoginAccountRemove
            Call WriteLoginRemove
            
        Case E_MODO.e_DisconnectForced
            Call WriteDisconnectForced
    End Select

End Sub
