Attribute VB_Name = "mEngine"
Option Explicit

Private Type FunctionInfo
    Name As String
    Address As Long
    Size As Long
    Checksum As Long
End Type

Private mFunctionList() As FunctionInfo

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function EnumResourceNamesA Lib "kernel32" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32.dll" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const TIMER_ID As Long = 1
Private Const TIMER_INTERVAL As Long = 3600000 ' 60 minutos en milisegundos

Private mModuleHandle As Long
Private mFunctionCount As Integer

Function CalculateChecksum(ptr As Long, Size As Long) As Long
    Dim i As Long
    Dim sum As Long
    For i = 1 To Size
        sum = sum + ByteAt(ptr + i - 1)
    Next i
    CalculateChecksum = sum
End Function

Function ByteAt(ptr As Long) As Byte
    ByteAt = ReadMemoryByte(ptr)
End Function

Private Function ReadMemoryByte(ByVal lpBaseAddress As Long) As Byte
    Dim lpBuffer As Byte
    Call CopyMemory(ByVal VarPtr(lpBuffer), ByVal lpBaseAddress, 1)
    ReadMemoryByte = lpBuffer
End Function

Private Sub CopyMemory(lpTo As Long, lpFrom As Long, ByVal cBytes As Long)
    Dim i As Long
    For i = 1 To cBytes
        Poke lpTo + i - 1, Peek(lpFrom + i - 1)
    Next i
End Sub

Private Sub Poke(ByVal Address As Long, ByVal Value As Variant)
    Call CopyMemory(Address, VarPtr(Value), Len(Value))
End Sub

Private Function Peek(ByVal Address As Long, Optional ByVal Size As Integer = 4) As Variant
    Dim Value As Variant
    Call CopyMemory(VarPtr(Value), Address, Size)
    Peek = Value
End Function

Private Function EnumerateFunction(ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal lParam As Long) As Long
    ' lpName contiene la dirección de la cadena que representa el nombre de la función
    Dim functionName As String
    functionName = Space$(256)
    Call CopyMemory(ByVal functionName, ByVal lpName, 256)

    ' Imprimir el nombre de la función
    Debug.Print Trim$(functionName)

    ' Obtener la dirección de la función
    Dim functionAddress As Long
    functionAddress = GetProcAddress(mModuleHandle, Trim$(functionName))

    ' Calcular el tamaño de la función (a modo de ejemplo, puedes ajustar según tus necesidades)
    Dim functionSize As Long
    functionSize = 4 ' Tamaño en bytes, ajustar según la función

    ' Calcular el checksum de la función
    Dim functionChecksum As Long
    functionChecksum = CalculateChecksum(functionAddress, functionSize)

    ' Aumentar la cantidad de funciones
    mFunctionCount = mFunctionCount + 1

    ' Redimensionar el array
    ReDim Preserve mFunctionList(1 To mFunctionCount)

    ' Almacenar la información de la función en el array
    mFunctionList(mFunctionCount).Name = Trim$(functionName)
    mFunctionList(mFunctionCount).Address = functionAddress
    mFunctionList(mFunctionCount).Size = functionSize
    mFunctionList(mFunctionCount).Checksum = functionChecksum

    ' Devolver 1 para continuar enumerando
    EnumerateFunction = 1
End Function

Public Sub VerifyIntegrity()
    Dim i As Integer
    For i = 1 To mFunctionCount
        ' Proteger la memoria para permitir escritura
        Dim oldProtect As Long
        VirtualProtect ByVal mFunctionList(i).Address, mFunctionList(i).Size, PAGE_EXECUTE_READWRITE, oldProtect

        ' Calcular el checksum actual de la función en memoria
        Dim currentChecksum As Long
        currentChecksum = CalculateChecksum(mFunctionList(i).Address, mFunctionList(i).Size)

        ' Restaurar la protección de la memoria original
        VirtualProtect ByVal mFunctionList(i).Address, mFunctionList(i).Size, oldProtect, oldProtect

        ' Comparar el checksum calculado con el almacenado
        If currentChecksum <> mFunctionList(i).Checksum Then
            MsgBox "¡Alerta! La función " & mFunctionList(i).Name & " ha sido modificada."
        End If
    Next i
End Sub
