Attribute VB_Name = "modNetwork"
Option Explicit

Public Reader As Network.Reader ' @NOTE: need refactoring, this is so its easier to use in every handler
Public Writer As Network.Writer ' @NOTE: same as above
Public Client As Network.Client

Public LogearCuenta As Boolean
Public IsConnected As Boolean
Public ConnectedIP As String
Public ConnectedPort As String

Private Const CLIENT_XOR_KEY As Long = 192
Private Const SERVER_XOR_KEY As Long = 128

Public Sub OnNetworkConnect()
    'FrmConnect.imgConnect.Enabled = True
    'FrmConnect.tEnabled.Enabled = False
    IsConnected = True
    
    FrmConectando.visible = False
    Call Writer.Clear
End Sub

Private Sub CloseOnReport(ByVal code As Long)
    On Error GoTo CloseOnReport_Err

    If (error = 10061) Then
      '  Call MsgBox("Asegúrate de estar conectado a internet y de que el servidor esté online desde https://www.argentumgame.com")
    Else
        If (error <> 0 And error <> 2) Then
          ' Call MsgBox("Asegúrate de estar conectado a internet y de que el servidor esté online desde https://www.argentumgame.com")
        End If
    End If
    
    Exit Sub

CloseOnReport_Err:

End Sub
Public Sub OnNetworkClose(ByVal code As Long)
    
    FrmConectando.visible = True
    IsConnected = False
    Call ResetAllInfo(True)
    Call CloseOnReport(code)
    
    'If TempAccount.Email = vbNullString Then
        
        FrmConectando.tReconnect.Enabled = True
        'FrmConectando.tReconnect.Interval = 2000
        'FrmConectando.Reconnect_Socket
   ' Else
        
       ' Account.Email = TempAccount.Email
      '  Account.Passwd = TempAccount.Passwd
    
      '  Prepare_And_Connect E_MODO.e_LoginAccount

   ' End If
End Sub

Public Sub OnNetworkSend(ByVal Message As Network.Reader)
'asdasdasd
  '  Dim BufferRef() As Byte
   ' Call Message.GetData(BufferRef)
    
 '   Dim i As Long
    'For i = 0 To UBound(BufferRef)
    '    BufferRef(i) = BufferRef(i) Xor CLIENT_XOR_KEY
   ' Next i
End Sub

Public Sub OnNetworkRecv(ByVal Message As Network.Reader)

    'Dim BufferRef() As Byte
    'Call Message.GetData(BufferRef)
    
   ' Dim i As Long
   ' For i = 0 To UBound(BufferRef)
      '  BufferRef(i) = BufferRef(i) Xor SERVER_XOR_KEY
    'Next i
    
    Set Reader = Message
    
    While (Message.GetAvailable > 0)
        Protocol.HandleIncomingData
    Wend
    
    Set Reader = Nothing
End Sub

Public Sub Initialise()
    Set Client = New Network.Client
    Set Writer = New Network.Writer
    
    Call Client.Attach(AddressOf OnNetworkConnect, AddressOf OnNetworkClose, AddressOf OnNetworkSend, AddressOf OnNetworkRecv)
End Sub

Public Sub Connect(ByVal Address As String, ByVal Service As String)

    ConnectedIP = Address
    ConnectedPort = Service
    
    Call Initialise
    Call Client.Connect(Address, Service)

End Sub

Public Sub Disconnect()
    If Not Client Is Nothing Then
        Call Client.Close(True)
    End If
End Sub

Public Sub Poll()
    If (Client Is Nothing) Then
        Exit Sub
    End If
    
    Call Client.Flush
    Call Client.Poll
End Sub

Public Sub Send(ByVal Urgent As Boolean)

    Call Client.Send(Urgent, Writer)
    Call Writer.Clear
End Sub
