Attribute VB_Name = "mServerLocal"
Option Explicit



Public ServerArchive As NETWORK.Server
Public WriterArchive As NETWORK.Writer
Public ReaderArchive As NETWORK.Reader

Private Const CLIENT_XOR_KEY As Long = 192
Private Const SERVER_XOR_KEY As Long = 128

Private Enum ClientPacketID_Archive
    Connected = 1
End Enum

Private Enum ServerPacketID_Archive
    Logged = 1
    LogSecurity = 2   ' Enviamos un LOG
End Enum


' Tipos de LOGS
Public Enum eLogSecurity_SubType
    sGeneral = 1
    sAntiFrag = 2
    sAntiCheat = 3
    sAntiFraude = 4
End Enum
Public Sub SocketConfig_Archive()

    On Error Resume Next

    Set WriterArchive = New NETWORK.Writer
    Set ServerArchive = New NETWORK.Server
    
    Call ServerArchive.Attach(AddressOf OnServerConnect_Archive, AddressOf OnServerClose_Archive, AddressOf OnServerSend_Archive, AddressOf OnServerReceive_Archive)
    Call ServerArchive.Listen(1, "0.0.0.0", 12501)
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
End Sub
Public Sub OnServerSend_Archive(ByVal Connection As Long, ByVal message As NETWORK.Reader)

    'Debug.Print "OnServerSend"
    
    Dim BufferRef() As Byte
    Call message.GetData(BufferRef)
    
    Dim i As Long
    For i = 0 To UBound(BufferRef)
        BufferRef(i) = BufferRef(i) Xor SERVER_XOR_KEY
    Next i
End Sub
Public Sub OnServerConnect_Archive(ByVal Connection As Long, ByVal Address As Long)
    
    If Connection <= 1 Then
       ' Call WriteLogSecurity(Connection, eAntiFrag, "Macro Externo", "Lautaro", "Melkor")
        'Call WriteConnectedMessage(Connection)
    Else
        Debug.Print ("No conection")
        Call Protocol.Kick(Connection, "El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
    End If
    
End Sub
Public Sub OnServerClose_Archive(ByVal Connection As Long)
On Error GoTo OnServerClose_Error
    
    Exit Sub

OnServerClose_Error:

    Call LogError("OnServerClose: " + Err.description)
    
End Sub

Public Sub OnServerReceive_Archive(ByVal Connection As Long, ByVal message As NETWORK.Reader)

    Dim BufferRef() As Byte
    Call message.GetData(BufferRef)
    
    Dim i As Long
    
    For i = 0 To UBound(BufferRef)
        BufferRef(i) = BufferRef(i) Xor CLIENT_XOR_KEY
    Next i
    
    Set ReaderArchive = message
    
    While (message.GetAvailable() > 0)

        Call HandleIncomingData_Archive(Connection)

    Wend
    
    Set ReaderArchive = Nothing
End Sub

Public Function HandleIncomingData_Archive(ByVal Connection As Integer) As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    '
    '***************************************************
    On Error Resume Next

    Dim PacketID As Byte
    
    Select Case PacketID
    
    End Select
   
    'Done with this packet, move on to next one or send everything if no more packets found
    If Err.Number = 0 Then
        Err.Clear
        
        HandleIncomingData_Archive = True
    ElseIf Err.Number <> 0 Then
        Call ServerArchive.Kick(Connection)
        HandleIncomingData_Archive = False
    Else
        HandleIncomingData_Archive = False
    End If
    
End Function


'############################################################ PACKETS ##########################################
Public Sub WriteConnectedMessage(ByVal Connection As Integer)

    Call WriterArchive.WriteInt(ServerPacketID_Archive.Logged)
    Call ServerArchive.Send(Connection, False, WriterArchive)

End Sub

Public Sub WriteLogSecurity(ByVal Connection As Integer, _
                            ByRef SubType As eLogSecurity_SubType, _
                            ByVal Argument As String, _
                            ByVal Responsable As String, _
                            ByVal Victima As String)

    Call WriterArchive.WriteInt8(ServerPacketID_Archive.LogSecurity)
    Call WriterArchive.WriteInt8(SubType)
    
    Call WriterArchive.WriteString8(Argument)
    Call WriterArchive.WriteString8(Responsable)
    Call WriterArchive.WriteString8(Victima)

    Call ServerArchive.Send(Connection, False, WriterArchive)

End Sub
