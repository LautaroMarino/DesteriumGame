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
        intXOrValue1 = Val("&H" & (mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(mid$(XOR_CHARACTER, ((lonDataPtr Mod Len(XOR_CHARACTER)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XORDecryption = strDataOut
End Function
