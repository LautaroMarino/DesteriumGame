Attribute VB_Name = "Application"
'**************************************************************
' Application.bas - General API methods regarding the Application in general.
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

''
' Retrieves the active window's hWnd for this app.
'
' @return Retrieves the active window's hWnd for this app. If this app is not in the foreground it returns 0.

Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Type tAccountSecurity
    IP_Public As String
    IP_Local As String
    
    SERIAL_MAC As String
    SERIAL_DISK As String
    SERIAL_MOTHERBOARD As String
    SERIAL_BIOS As String
    SERIAL_PROCESSOR As String
     
    SYSTEM_DATA As String
End Type

Public AccountSec As tAccountSecurity


' Obtiene datos básicos del sistema (Sistema operativo, memoria ram y procesador)
Public Function System_GetData() As String
    On Error GoTo ErrHandler
    
        Dim Results As Object, Info As Object, PCInfo As String, Ram As String, TotMem As Long
    
        ' Get the Memory information. For more information from this query, see: https://msdn.microsoft.com/en-us/library/aa394347(v=vs.85).aspx
        Set Results = GetObject("Winmgmts:").ExecQuery("SELECT Capacity FROM Win32_PhysicalMemory")
        For Each Info In Results
            'Capacity returns the size separately for each stick in bytes. Therefore we loop and add whilst dividing by 1GB.
            TotMem = TotMem + (Info.Capacity / 1073741824)
        Next Info
    
        ' Get the O.S. information. For more information from this query, see: https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
        Set Results = GetObject("Winmgmts:").ExecQuery("SELECT Caption,Version,ServicePackMajorVersion,ServicePackMinorVersion,OSArchitecture,TotalVisibleMemorySize FROM Win32_OperatingSystem")
        For Each Info In Results 'Info.Version can be used to calculate Windows version. E.G. If Val(Left$(Info.Version,3)>=6.1 then it is at least Windows 7.
            PCInfo = Info.Caption & " - " & Info.version & "  SP " & _
                Info.ServicePackMajorVersion & "." & _
                Info.ServicePackMinorVersion & "  " & Info.OSArchitecture & "  " & vbNewLine
            Ram = "Installed RAM: " & Format$(TotMem, "0.00 GB (") & Format$(Info.TotalVisibleMemorySize / 1048576, "0.00 GB usable)") 'Divide by 1MB to get GB
        Next Info
    
        ' Get the C.P.U. information. For more information from this query, see: https://msdn.microsoft.com/en-us/library/aa394373(v=vs.85).aspx
        Set Results = GetObject("Winmgmts:").ExecQuery("SELECT Name,AddressWidth,NumberOfLogicalProcessors,CurrentClockSpeed FROM Win32_Processor")
        For Each Info In Results
            PCInfo = PCInfo & Info.Name & "  " & Info.AddressWidth & _
                "-bit." & vbNewLine & Info.NumberOfLogicalProcessors & _
                " Cores " & Info.CurrentClockSpeed & "MHz.  " & Ram
        Next Info
    
        
        Set Results = Nothing
        System_GetData = PCInfo
    
    Exit Function
ErrHandler:
    Set Results = Nothing
    
End Function

' Obtiene la IP local de la LAN (Evitar doble clientes)
Public Function System_GetIP_Local() As String
    On Error GoTo ErrHandler
    
        Dim IPConfig As Variant
        Dim IPConfigSet As Object
        
        Set IPConfigSet = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
        
        For Each IPConfig In IPConfigSet
            If Not IsNull(IPConfig.IpAddress) Then
                If System_IP_Valid(IPConfig.IpAddress(0)) Then
                   System_GetIP_Local = IPConfig.IpAddress(0)
                End If
            End If
        Next IPConfig
    Exit Function
ErrHandler:
End Function

' Comprueba si la IP ingresada tiene formato '255.255.255.255'
Private Function System_IP_Valid(IpAddress As String) As Boolean

    Dim A As String, B As String, C As String, D As String
    
    ' No IP address is over 15 digits long inculding the "."
    If Len(IpAddress) > 15 Then System_IP_Valid = False: Exit Function
    A = InStr(1, IpAddress, ".", vbTextCompare)
    If A > 4 Then System_IP_Valid = False: Exit Function
    If CStr(mid$(IpAddress, 1, A)) > 255 Then System_IP_Valid = False: Exit Function
    B = InStr(A + 1, IpAddress, ".", vbTextCompare)
    If B - A > 4 Then System_IP_Valid = False: Exit Function
    If CStr(mid$(IpAddress, A + 1, B - A)) > 255 Then System_IP_Valid = False: Exit Function
    C = InStr(B + 1, IpAddress, ".", vbTextCompare)
    If C - B > 4 Then System_IP_Valid = False: Exit Function
    If CStr(mid$(IpAddress, B + 1, C - B)) > 255 Then System_IP_Valid = False: Exit Function
    D = Len(mid$(IpAddress, C + 1, Len(IpAddress)))
    If D - C > 3 Then System_IP_Valid = False: Exit Function
    If CStr(mid$(IpAddress, C + 1, D)) > 255 Then System_IP_Valid = False: Exit Function
    System_IP_Valid = True
End Function

' Obtiene el MAC Adress
Public Function System_GetMAC()
    On Error GoTo ErrHandler

        Dim colNetAdapters, objWMIService As Object, Temp
        Dim strComputer As String, objItem As Object
        
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
        
        For Each objItem In colNetAdapters
            Temp = objItem.MACAddress
        Next
        
        System_GetMAC = Temp
    Exit Function
ErrHandler:
End Function

' Obtiene el SERIAL MOTHERBOARD
Public Function System_GetSerial_Motherboard() As String
On Error GoTo ErrHandler
    
    Dim List, Object, Temp
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BaseBoard")
    
    For Each Object In List
        Temp = Object.SerialNumber
    Next
    
    System_GetSerial_Motherboard = Temp
   Exit Function
ErrHandler:
   
End Function

' Obtiene el SERIAL PROCESADOR
Public Function System_GetSerial_Processor() As String
On Error GoTo ErrHandler
    
    Dim List, Object, Temp
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    
    For Each Object In List
        Temp = Object.UniqueID
    Next
   
   If Not Temp = vbNull Then
        System_GetSerial_Processor = Temp
   End If
    
    
    Exit Function
ErrHandler:
   
End Function

' Obtiene el SERIAL BIOS
Public Function System_GetSerial_BIOS() As String
On Error GoTo ErrHandler
    
    Dim Temp
    Dim List, Object
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BIOS")
    
    For Each Object In List
        Temp = Object.SerialNumber
    Next
    
    System_GetSerial_BIOS = Temp
    
   Exit Function
ErrHandler:
End Function

' Obtiene el SERIAL DISCO
Public Function System_GetSerial_DISK() As String
On Error GoTo ErrHandler
    
    Dim Temp
    Dim List, Object
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDisk")
    
    For Each Object In List
        Temp = Object.VolumeSerialNumber
        System_GetSerial_DISK = Temp
        Exit Function
    Next
    
   Exit Function
ErrHandler:
End Function

' Obtiene la IP pública de la web
Public Function IP_Publica() As String

    On Error GoTo ErrHandler

    Dim cTemp    As String

    Dim arTemp() As String

    Dim url      As String

    Dim IP       As String

    IP = "c:\ip.txt"
    'URL = "http://miip.es"
    url = "http://miip.es"

    'URL = "http://myip.es" 'AUN NO FUNCIONA
    If Dir(IP) <> "" Then Kill IP
    Call URLDownloadToFile(0, url, IP, 0, 0)

    If Dir(IP) <> "" Then
        cTemp = CreateObject("Scripting.FileSystemObject").OpenTextFile(IP).ReadAll

        If url = "http://miip.es" Then
            If InStr(cTemp, "<h2>") > 0 Then
                arTemp = Split(Replace(cTemp, "</h2>", "<h2>"), "<h2>")
                IP_Publica = Trim(Right(arTemp(1), Len(arTemp(1)) - 9))
            End If

        ElseIf url = "http://www.cualesmiip.com" Then

            If InStr(cTemp, "Cual es mi IP Tu IP real es ") > 0 Then
                arTemp = Split(Replace(cTemp, " (", "Cual es mi IP Tu IP real es "), "Cual es mi IP Tu IP real es ")
                IP_Publica = Trim(arTemp(1))
            End If

        ElseIf url = "http://myip.es" Then

            If InStr(cTemp, "<aside>") > 0 Then
                arTemp = Split(Replace(cTemp, "</aside>", "<aside>"), "<aside>")
                IP_Publica = Trim(Right(arTemp(1), Len(arTemp(1)) - 9))
            End If
        End If

        Kill IP
    End If
    
    Exit Function

ErrHandler:

End Function














''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    IsAppActive = (GetActiveWindow <> 0)
End Function

