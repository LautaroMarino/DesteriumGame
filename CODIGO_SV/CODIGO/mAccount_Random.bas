Attribute VB_Name = "mAccount_Random"
Option Explicit

Public Const C_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890abcdefghijklmnopqrstuvwxyz"

Public Palabras_Profesiones() As String

Public Palabras_Alimentos()   As String

Public Palabras_Colores()     As String

Public Palabras_Transporte()  As String

Public Palabras_Geometricas() As String

Public Function RandomKey_Generate(Optional ByVal Lenght As Byte = 10, _
                                   Optional ByVal Characters As String = C_CHARACTERS) As String
        '<EhHeader>
        On Error GoTo RandomKey_Generate_Err
        '</EhHeader>

        Dim i        As Integer

        Dim Longitud As Integer
    
100     Longitud = Len(Characters)
    
102     For i = 1 To Lenght
104         RandomKey_Generate = RandomKey_Generate & mid(Characters, Int((Longitud * Rnd) + 1), 1)
106     Next i

        '<EhFooter>
        Exit Function

RandomKey_Generate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAccount_Random.RandomKey_Generate " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function RandomPasswd_Generate() As String
        '<EhHeader>
        On Error GoTo RandomPasswd_Generate_Err
        '</EhHeader>

        Dim A      As Long, B As Long

        Dim Random As Byte

        Dim Temp   As String
    
100     Random = RandomNumber(1, 100)
    
102     If Random <= 20 Then
104         Temp = Palabras_Profesiones(RandomNumber(LBound(Palabras_Profesiones), UBound(Palabras_Profesiones)))
106         Temp = Temp & RandomNumber(1, 100) & Palabras_Alimentos(RandomNumber(LBound(Palabras_Alimentos), UBound(Palabras_Alimentos)))
108         RandomPasswd_Generate = Temp

            Exit Function

110     ElseIf Random <= 40 Then
112         Temp = Palabras_Alimentos(RandomNumber(LBound(Palabras_Alimentos), UBound(Palabras_Alimentos)))
114         Temp = Temp & RandomNumber(100, 200) & Palabras_Colores(RandomNumber(LBound(Palabras_Colores), UBound(Palabras_Colores)))
116         RandomPasswd_Generate = Temp

            Exit Function

118     ElseIf Random <= 60 Then
120         Temp = Palabras_Colores(RandomNumber(LBound(Palabras_Colores), UBound(Palabras_Colores)))
122         Temp = Temp & RandomNumber(200, 300) & Palabras_Transporte(RandomNumber(LBound(Palabras_Transporte), UBound(Palabras_Transporte)))
124         RandomPasswd_Generate = Temp

            Exit Function

126     ElseIf Random <= 80 Then
128         Temp = Palabras_Transporte(RandomNumber(LBound(Palabras_Transporte), UBound(Palabras_Transporte)))
130         Temp = Temp & RandomNumber(300, 400) & Palabras_Geometricas(RandomNumber(LBound(Palabras_Geometricas), UBound(Palabras_Geometricas)))
132         RandomPasswd_Generate = Temp

            Exit Function

134     ElseIf Random <= 100 Then
136         Temp = Palabras_Transporte(RandomNumber(LBound(Palabras_Transporte), UBound(Palabras_Transporte)))
138         Temp = Temp & RandomNumber(400, 500) & Palabras_Alimentos(RandomNumber(LBound(Palabras_Alimentos), UBound(Palabras_Alimentos)))
140         RandomPasswd_Generate = Temp

            Exit Function

        End If
    
        '<EhFooter>
        Exit Function

RandomPasswd_Generate_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAccount_Random.RandomPasswd_Generate " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Se cargan las palabras aleatoreas
Public Sub LoadPalabras_All()
    
    Call LoadPalabras("Palabras_Profesiones.txt", Palabras_Profesiones)
    Call LoadPalabras("Palabras_Alimentos.txt", Palabras_Alimentos)
    Call LoadPalabras("Palabras_Colores.txt", Palabras_Colores)
    Call LoadPalabras("Palabras_Transporte.txt", Palabras_Transporte)
    Call LoadPalabras("Palabras_Geometricas.txt", Palabras_Geometricas)
    
End Sub

' Funciones generales de escritura única.
'############################################
'############################################

Public Sub LoadPalabras(ByVal FilePath As String, _
   ByRef Arrai() As String)
        '<EhHeader>
        On Error GoTo LoadPalabras_Err
        '</EhHeader>
                        
        Dim File        As Long

        Dim Tmp         As String

        Dim A           As Long

        Dim CopyArray() As String
    
100     FilePath = DatPath & FilePath
102     File = FreeFile()
    
104     Open FilePath For Input As #File
    
106     Do While Not EOF(File)
108         ReDim Preserve CopyArray(0 To A) As String
110         Line Input #File, Tmp
112         CopyArray(A) = LCase$(Tmp)
114         A = A + 1
        Loop
    
116     Close #File
    
118     Arrai = CopyArray
        '<EhFooter>
        Exit Sub

LoadPalabras_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mAccount_Random.LoadPalabras " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

