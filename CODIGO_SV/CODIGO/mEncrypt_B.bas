Attribute VB_Name = "mEncrypt_B"
Option Explicit

Public Const XOR_CHARACTER As String = "6vmEL7VdvNkFi0iY6WLh"

'Codigo fuente: https://www.freevbcode.com/ShowCode.asp?ID=5676
Public Function XORDecryption(DataIn As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = val("&H" & (mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(mid$(XOR_CHARACTER, ((lonDataPtr Mod Len(XOR_CHARACTER)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XORDecryption = strDataOut
End Function


Public Function XOREncryption(ByVal Texto As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim Temp As Integer
    Dim tempstring As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To Len(Texto)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Asc(mid$(Texto, lonDataPtr, 1))
        'The second value comes from the code key
        intXOrValue2 = Asc(mid$(XOR_CHARACTER, ((lonDataPtr Mod Len(XOR_CHARACTER)) + 1), 1))
        
        Temp = (intXOrValue1 Xor intXOrValue2)
        tempstring = hex(Temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        
        strDataOut = strDataOut + tempstring
    Next lonDataPtr
   XOREncryption = strDataOut
End Function


' Busca informacion del usuario
Public Sub Security_SearchData(ByVal UserIndex As Integer, ByVal Selected As eSearchData, ByVal Data As String)
    Dim A As Long
    Dim Users As String
    
    For A = 1 To LastUser
        If UserList(A).flags.UserLogged Then
            Select Case Selected
                Case eSearchData.eMac
                    If StrComp(UserList(A).Account.Sec.SERIAL_MAC, Data) = 0 Then
                        Users = Users & UserList(A).Name & ", "
                    End If
                Case eSearchData.eDisk
                    If StrComp(UserList(A).Account.Sec.SERIAL_DISK, Data) = 0 Then
                        Users = Users & UserList(A).Name & ", "
                    End If
                Case eSearchData.eIpAddress
                    If UserList(A).Account.Sec.IP_Address = CLng(Data) Then
                        Users = Users & UserList(A).Name & ", "
                    End If
                Case Else
            End Select
        End If
    Next A
    
    If Len(Users) > 2 Then
        Users = Left$(Users, Len(Users) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios encontrados: " & Users & ".", FontTypeNames.FONTTYPE_INFOGREEN)
    Else
        Call WriteConsoleMsg(UserIndex, "No se han encontrado usuarios en la búsqueda.", FontTypeNames.FONTTYPE_INFORED)
    End If
    
End Sub




