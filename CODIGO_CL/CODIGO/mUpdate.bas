Attribute VB_Name = "mUpdate"
Option Explicit

Public Sub UpdateMd5File()
    On Error GoTo ErrHandler
    
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim i As Integer

    ' Especifica la ruta completa del archivo Md5.txt
    filePath = App.path & "\Md5Classic.txt"

    ' Verifica si el archivo existe
    If FileExists(filePath) Then
        ' Lee el contenido del archivo
        Open filePath For Input As #1
        fileContent = Input$(LOF(1), #1)
        Close #1

        ' Divide el contenido en líneas
        lines = Split(fileContent, vbLf)

        ' Modifica las líneas específicas
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "DesteriumHD.exe") > 0 Or InStr(lines(i), "DesteriumClassic.exe") > 0 Then
                ' Encuentra las líneas que contienen los nombres de los archivos y actualiza el hash
                lines(i) = UpdateHash(lines(i))
            End If
        Next i

        ' Une las líneas modificadas
        fileContent = Join(lines, vbLf)

        ' Guarda los cambios en el archivo
        Open filePath For Output As #1
        Print #1, fileContent;
        Close #1
        
        MsgBox "El cliente se cerrará por actualización obligatoria. Entra nuevamente y lograrás entrar correctamente.", vbInformation
        prgRun = False
    End If
    
ErrHandler:
    Exit Sub
    
End Sub

Private Function UpdateHash(line As String) As String
    ' Actualiza el hash agregando la fecha en formato numérico (ejemplo: "hash existente" -> "hash existente 17112023")
    Dim parts() As String
    parts = Split(line, "-")

    If UBound(parts) > 0 And InStr(line, "manifest") = 0 Then
        ' No modifica las líneas que contienen "MANIFEST"
        
        ' Obtén la fecha en formato numérico
        Dim numericDate As String
        numericDate = Format(Now, "ddmmyyyy")

        ' Agrega la fecha al final del hash existente
        UpdateHash = parts(0) & "-" & parts(1) & numericDate
    Else
        ' Si no hay un hash existente o la línea contiene "MANIFEST", devuelve la línea sin cambios
        UpdateHash = line
    End If
End Function

Private Function FileExists(filePath As String) As Boolean
    ' Verifica si un archivo existe
    On Error Resume Next
    FileExists = (GetAttr(filePath) And vbDirectory) = 0
    On Error GoTo 0
End Function

Private Function MakeArray(ParamArray args() As Variant) As Variant
    ' Función para crear un array a partir de una lista de argumentos
    MakeArray = args
End Function
